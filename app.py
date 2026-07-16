"""
app.py - PILLAR GloBE/DAC9
Flusso: apri sito → Microsoft login automatico → home upload
Supporta upload multi-CF: ogni file XML può avere un CF fornitore differente.

Hardening applicato (Maggio 2026):
- Storage persistente (Azure Blob Storage con fallback locale per dev) per
  file caricati, output pipeline, e stato dei job — risolve la perdita dati
  dovuta al filesystem effimero di App Service F1.
- Logging strutturato su stream (raccolto da App Service log / App Insights
  se configurato) invece che solo su liste in RAM.
"""
import logging
import os
import tempfile
import threading
import time
import uuid
from datetime import timedelta
from functools import wraps
from pathlib import Path

import msal
from flask import Flask, redirect, render_template, request, send_file, session, url_for, jsonify, abort
from werkzeug.middleware.proxy_fix import ProxyFix

from services.storage import storage

# ── Logging ──────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=os.getenv("LOG_LEVEL", "INFO"),
    format="%(asctime)s %(levelname)s %(name)s :: %(message)s",
)
log = logging.getLogger("pillar.app")

# ── Config ──────────────────────────────────────────────────────────────────
DEFAULT_CF = os.getenv("PILLAR_CF", "00731410155")

CLIENT_ID     = os.getenv("AZURE_AD_CLIENT_ID", "")
CLIENT_SECRET = os.getenv("AZURE_AD_CLIENT_SECRET", "")
TENANT_ID     = os.getenv("AZURE_TENANT_ID", "")
AUTHORITY     = "https://login.microsoftonline.com/" + TENANT_ID
SCOPE         = ["User.Read"]

app = Flask(__name__)
app.wsgi_app = ProxyFix(app.wsgi_app, x_proto=1, x_host=1)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "")
if not app.secret_key:
    # ATTENZIONE: senza FLASK_SECRET_KEY fisso, ogni riavvio invalida tutte
    # le sessioni attive (bug noto su F1 dove i riavvii sono frequenti).
    # Impostare sempre FLASK_SECRET_KEY come App Setting (idealmente da
    # Key Vault) in produzione.
    app.secret_key = os.urandom(32).hex()
    log.warning("FLASK_SECRET_KEY non impostata: generata a runtime "
                "(le sessioni non sopravvivono a un riavvio). Impostarla "
                "come App Setting persistente in produzione.")
app.permanent_session_lifetime = timedelta(days=7)

# ── Auth helpers ─────────────────────────────────────────────────────────────

def _msal_app():
    return msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET,
    )

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get("user"):
            if CLIENT_ID:
                session["next"] = request.url
                flow = _msal_app().initiate_auth_code_flow(
                    SCOPE,
                    redirect_uri=url_for("auth_callback", _external=True),
                )
                session["flow"] = flow
                return redirect(flow["auth_uri"])
            else:
                session["user"] = {"name": "Dev Locale", "email": "dev@localhost"}
        return f(*args, **kwargs)
    return decorated

@app.context_processor
def inject_user():
    return {"current_user": session.get("user")}

# ── Auth routes ───────────────────────────────────────────────────────────────

@app.route("/auth/callback")
def auth_callback():
    flow = session.pop("flow", None)
    if not flow:
        return redirect(url_for("index"))
    try:
        result = _msal_app().acquire_token_by_auth_code_flow(flow, request.args)
    except Exception:
        log.exception("Errore durante acquire_token_by_auth_code_flow")
        return redirect(url_for("index"))

    if "error" in result:
        log.warning("Login fallito: %s", result.get("error"))
        return f"Errore login: {result.get('error_description', result['error'])}", 400

    claims = result.get("id_token_claims", {})

    # Blocca account consumer Microsoft personali
    if claims.get("tid") == "9188040d-6c67-4c5b-b112-36a304b66dad":
        log.warning("Accesso negato: account consumer")
        return "Accesso negato: usa un account Microsoft aziendale.", 403

    email = claims.get("preferred_username", "").lower()

    session["user"] = {
        "name":  claims.get("name", claims.get("preferred_username", "Utente")),
        "email": email,
    }
    session.permanent = True
    log.info("Login OK: %s", email)
    return redirect(session.pop("next", url_for("index")))

@app.route("/auth/logout")
def auth_logout():
    email = (session.get("user") or {}).get("email")
    session.clear()
    log.info("Logout: %s", email)
    logout_url = (
        f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/logout"
        f"?post_logout_redirect_uri={url_for('index', _external=True)}"
    )
    return redirect(logout_url if TENANT_ID else url_for("index"))

# ── Job store (persistente via services.storage) ──────────────────────────────
# In precedenza: dict in RAM (JOBS = {}), perso ad ogni riavvio del processo
# (frequente su F1). Ora ogni job è un blob JSON — sopravvive a riavvii,
# consultabile per audit/troubleshooting.

_JOB_LOCK = threading.Lock()  # serializza le scritture concorrenti sullo stesso job

def new_job(files):
    jid = str(uuid.uuid4())
    job = {
        "status": "PENDING", "files": files,
        "outputs": [], "log": [], "error": None,
        "t0": time.time(), "t1": None,
    }
    storage.save_job(jid, job)
    log.info("Job creato: %s (%d file)", jid, len(files))
    return jid

def update_job(jid, **kw):
    with _JOB_LOCK:
        job = storage.load_job(jid) or {}
        job.update(kw)
        storage.save_job(jid, job)

def log_job(jid, msg):
    with _JOB_LOCK:
        job = storage.load_job(jid) or {"log": []}
        job.setdefault("log", []).append(msg)
        storage.save_job(jid, job)
    log.info("[job %s] %s", jid, msg)

def get_job(jid):
    return storage.load_job(jid)

# ── Pipeline worker ───────────────────────────────────────────────────────────

def run_pipeline(jid, items):
    """
    items: list of (blob_name: str, cf: str)
    Ogni item viene scaricato in una cartella temporanea locale (necessaria
    perché lxml/openpyxl lavorano su filesystem), processato, e gli output
    vengono ricaricati sullo storage persistente. La cartella temporanea
    viene distrutta a fine job: nessun dato residuo sul filesystem effimero.

    NOTA (2026-07): import dei tre moduli e update_job(RUNNING) spostati
    dentro il try/except. Prima erano fuori: se l'import falliva (es. deploy
    Kudu non atomico che lascia temporaneamente un services/*.py incoerente,
    o un errore di sintassi introdotto per sbaglio), l'eccezione non veniva
    mai catturata, il thread daemon moriva silenziosamente e il job restava
    bloccato per sempre a PENDING/RUNNING — con il frontend in polling
    infinito ("resta in caricamento"). Ora qualsiasi fallimento, incluso
    quello a livello di import, marca il job FAILED con errore visibile.
    """
    outputs = []

    try:
        from services.xml_validator import validate
        from services.shell_telematico import encapsulate
        from services.report_viewer import generate_report

        update_job(jid, status="RUNNING")
        with tempfile.TemporaryDirectory(prefix=f"pillar_{jid}_") as tmp:
            tmp_dir = Path(tmp)

            for blob_name, cf in items:
                xp = tmp_dir / Path(blob_name).name
                storage.download_to(blob_name, xp)

                out_dir = tmp_dir / "out" / xp.stem
                out_dir.mkdir(parents=True, exist_ok=True)
                log_job(jid, f"▶ {xp.name}  (CF: {cf})")

                for label, fn, kw in [
                    ("Step 1 – Validazione",  validate,        {"output_dir": out_dir}),
                    ("Step 2 – Shell MSG",    encapsulate,     {"codice_fiscale": cf, "output_dir": out_dir, "archive_source": False}),
                    ("Step 3 – Report HTML",  generate_report, {"output_dir": out_dir}),
                ]:
                    try:
                        log_job(jid, f"  {label}...")
                        if label.startswith("Step 2"):
                            r = fn(xml_path=xp, **kw)
                        else:
                            r = fn(xp, **kw)

                        # upload dell'output verso lo storage persistente
                        out_blob = f"results/{jid}/{xp.stem}/{r.name}"
                        storage.save_file(r, out_blob)

                        outputs.append({
                            "label": f"[{xp.stem}] {label[9:]}",
                            "filename": r.name,
                            "blob": out_blob,
                        })
                        log_job(jid, f"  ✓ {r.name}")
                    except Exception as e:
                        log_job(jid, f"  ✗ {label}: {e}")
                        log.exception("Errore in %s per job %s", label, jid)

                log_job(jid, f"  ✓ Completato")

        update_job(jid, status="COMPLETED", outputs=outputs, t1=time.time())
        log.info("Job %s completato: %d output", jid, len(outputs))
    except Exception as e:
        update_job(jid, status="FAILED", error=str(e), t1=time.time())
        log_job(jid, f"ERRORE: {e}")
        log.exception("Job %s fallito", jid)

# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/")
@login_required
def index():
    return render_template("index.html", default_cf=DEFAULT_CF)

@app.route("/upload", methods=["POST"])
@login_required
def upload():
    files   = request.files.getlist("xml_files")
    cf_map  = request.form.getlist("cf_map")   # parallel list: cf_map[i] → files[i]

    if not files or all(f.filename == "" for f in files):
        return jsonify({"error": "Nessun file XML"}), 400

    items = []   # (blob_name, cf)
    saved_names = []

    with tempfile.TemporaryDirectory(prefix="pillar_upload_") as tmp:
        tmp_dir = Path(tmp)
        for i, f in enumerate(files):
            if not f.filename:
                continue
            name = Path(f.filename).name
            if not name.lower().endswith(".xml"):
                return jsonify({"error": f"Solo file .xml ({name})"}), 400

            cf = cf_map[i].strip().upper() if i < len(cf_map) else DEFAULT_CF
            if len(cf) < 10:
                return jsonify({"error": f"CF non valido per il file '{name}': '{cf}'"}), 400

            local_tmp = tmp_dir / name
            f.save(str(local_tmp))

            blob_name = f"uploads/{uuid.uuid4()}/{name}"
            storage.save_file(local_tmp, blob_name)

            items.append((blob_name, cf))
            saved_names.append(name)

    if not items:
        return jsonify({"error": "Nessun file valido"}), 400

    jid = new_job([{"name": n} for n in saved_names])
    threading.Thread(
        target=run_pipeline,
        args=(jid, items),
        daemon=True,
    ).start()
    return jsonify({"job_id": jid, "files": saved_names}), 202

@app.route("/status/<jid>")
@login_required
def status(jid):
    j = get_job(jid)
    if not j:
        return jsonify({"error": "Job non trovato"}), 404
    elapsed = round(j["t1"] - j["t0"], 1) if j.get("t1") else None
    return jsonify({
        "job_id": jid,
        "status": j["status"],
        "log": j["log"],
        "outputs": j["outputs"],
        "error": j["error"],
        "elapsed": elapsed,
    })

@app.route("/download/<jid>/<filename>")
@login_required
def download(jid, filename):
    j = get_job(jid)
    if not j:
        abort(404)
    target = next(
        (o["blob"] for o in j["outputs"] if o["filename"] == filename),
        None,
    )
    if not target or not storage.exists(target):
        abort(404)
    log.info("Download: job=%s file=%s user=%s", jid, filename,
              (session.get("user") or {}).get("email"))
    return send_file(
        storage.open_read(target),
        as_attachment=True,
        download_name=filename,
    )

if __name__ == "__main__":
    print(f"\n  PILLAR | Auth: {'MSAL Azure AD' if CLIENT_ID else 'BYPASS locale'} | http://127.0.0.1:5000\n")
    app.run(debug=True, host="0.0.0.0", port=5000)