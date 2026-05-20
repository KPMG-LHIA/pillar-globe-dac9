"""
app.py - PILLAR GloBE/DAC9
Flusso: apri sito → Microsoft login automatico → home upload
"""
import os, threading, time, traceback, uuid
from datetime import timedelta
from functools import wraps
from pathlib import Path

import msal
from flask import Flask, redirect, render_template, request, send_file, session, url_for, jsonify, abort
from werkzeug.middleware.proxy_fix import ProxyFix

# ── Config  ──────────────────────────────────────────────────────────────────
BASE_DIR    = Path(__file__).parent
UPLOAD_DIR  = BASE_DIR / "uploads"
RESULTS_DIR = BASE_DIR / "results"
DEFAULT_CF  = os.getenv("PILLAR_CF", "")

UPLOAD_DIR.mkdir(exist_ok=True)
RESULTS_DIR.mkdir(exist_ok=True)

CLIENT_ID     = os.getenv("AZURE_AD_CLIENT_ID", "")
CLIENT_SECRET = os.getenv("AZURE_AD_CLIENT_SECRET", "")
TENANT_ID     = os.getenv("AZURE_TENANT_ID", "")
AUTHORITY     = "https://login.microsoftonline.com/organizations"  # multitenant: accetta qualsiasi tenant aziendale
SCOPE         = ["User.Read"]

app = Flask(__name__)
app.wsgi_app = ProxyFix(app.wsgi_app, x_proto=1, x_host=1)
app.secret_key = os.getenv("FLASK_SECRET_KEY", os.urandom(32).hex())
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
            # Nessuna pagina intermedia — redirect diretto a Microsoft
            if CLIENT_ID:
                session["next"] = request.url
                flow = _msal_app().initiate_auth_code_flow(
                    SCOPE,
                    redirect_uri=url_for("auth_callback", _external=True),
                )
                session["flow"] = flow
                return redirect(flow["auth_uri"])
            else:
                # Sviluppo locale senza credenziali: bypass
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
        return redirect(url_for("index"))

    if "error" in result:
        return f"Errore login: {result.get('error_description', result['error'])}", 400

    claims = result.get("id_token_claims", {})

    # Blocca account consumer Microsoft personali
    if claims.get("tid") == "9188040d-6c67-4c5b-b112-36a304b66dad":
        return "Accesso negato: usa un account Microsoft aziendale.", 403

    email = claims.get("preferred_username", "").lower()

    # Verifica dominio: accetta solo account @kpmg.* (es. @kpmg.it, @kpmg.com, @kpmg.de)
    # Il dominio deve iniziare con "kpmg." per evitare falsi positivi (es. notkpmg.com)
    domain = email.split("@")[-1] if "@" in email else ""
    if not (domain == "kpmg" or domain.startswith("kpmg.")):
        return (
            f"<h2 style='font-family:sans-serif;margin:40px'>Accesso negato</h2>"
            f"<p style='font-family:sans-serif;margin:40px'>L'account <strong>{email}</strong> "
            f"non appartiene al dominio KPMG. Solo account @kpmg.* sono autorizzati.</p>"
            f"<p style='font-family:sans-serif;margin:40px'><a href='/auth/logout'>Torna al login</a></p>"
        ), 403

    session["user"] = {
        "name":  claims.get("name", claims.get("preferred_username", "Utente")),
        "email": email,
    }
    session.permanent = True
    return redirect(session.pop("next", url_for("index")))

@app.route("/auth/logout")
def auth_logout():
    session.clear()
    logout_url = (
        f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/logout"
        f"?post_logout_redirect_uri={url_for('index', _external=True)}"
    )
    return redirect(logout_url if TENANT_ID else url_for("index"))

# ── Job store ─────────────────────────────────────────────────────────────────

JOBS = {}
LOCK = threading.Lock()

def new_job(files):
    jid = str(uuid.uuid4())
    with LOCK:
        JOBS[jid] = {"status": "PENDING", "files": files, "outputs": [], "log": [], "error": None, "t0": time.time(), "t1": None}
    return jid

def update_job(jid, **kw):
    with LOCK: JOBS[jid].update(kw)

def log_job(jid, msg):
    with LOCK: JOBS[jid]["log"].append(msg)

def get_job(jid):
    with LOCK: return JOBS.get(jid)

# ── Pipeline worker ───────────────────────────────────────────────────────────

def run_pipeline(jid, xml_paths, cf):
    from services.xml_validator import validate
    from services.shell_telematico import encapsulate
    from services.report_viewer import generate_report

    update_job(jid, status="RUNNING")
    outputs = []

    try:
        for xp in xml_paths:
            out_dir = RESULTS_DIR / xp.stem
            out_dir.mkdir(parents=True, exist_ok=True)
            log_job(jid, f"▶ {xp.name}")

            for label, fn, kw in [
                ("Step 1 – Validazione",    validate,        {"output_dir": out_dir}),
                ("Step 2 – Shell MSG",      encapsulate,     {"codice_fiscale": cf, "output_dir": out_dir, "archive_source": False}),
                ("Step 3 – Report HTML",    generate_report, {"output_dir": out_dir}),
            ]:
                try:
                    log_job(jid, f"  {label}...")
                    if label.startswith("Step 2"):
                        r = fn(xml_path=xp, **kw)
                    else:
                        r = fn(xp, **kw)
                    outputs.append({"label": f"[{xp.stem}] {label[9:]}", "filename": r.name, "path": str(r)})
                    log_job(jid, f"  ✓ {r.name}")
                except Exception as e:
                    log_job(jid, f"  ✗ {label}: {e}")

            log_job(jid, f"  ✓ Completato")

        update_job(jid, status="COMPLETED", outputs=outputs, t1=time.time())
    except Exception as e:
        update_job(jid, status="FAILED", error=str(e), t1=time.time())
        log_job(jid, f"ERRORE: {e}")

# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/")
@login_required
def index():
    return render_template("index.html", default_cf=DEFAULT_CF)

@app.route("/upload", methods=["POST"])
@login_required
def upload():
    files = request.files.getlist("xml_files")
    cf    = request.form.get("codice_fiscale", DEFAULT_CF).strip().upper()
    if not files or all(f.filename == "" for f in files):
        return jsonify({"error": "Nessun file XML"}), 400
    if not cf:
        return jsonify({"error": "Codice Fiscale obbligatorio"}), 400
    saved = []
    for f in files:
        if not f.filename: continue
        name = Path(f.filename).name
        if not name.lower().endswith(".xml"):
            return jsonify({"error": f"Solo file .xml ({name})"}), 400
        dest = UPLOAD_DIR / name
        f.save(str(dest))
        saved.append({"name": name, "path": str(dest)})
    if not saved:
        return jsonify({"error": "Nessun file valido"}), 400
    jid = new_job(saved)
    threading.Thread(target=run_pipeline, args=(jid, [Path(s["path"]) for s in saved], cf), daemon=True).start()
    return jsonify({"job_id": jid, "files": [s["name"] for s in saved]}), 202

@app.route("/status/<jid>")
@login_required
def status(jid):
    j = get_job(jid)
    if not j: return jsonify({"error": "Job non trovato"}), 404
    elapsed = round(j["t1"] - j["t0"], 1) if j["t1"] else None
    return jsonify({"job_id": jid, "status": j["status"], "log": j["log"], "outputs": j["outputs"], "error": j["error"], "elapsed": elapsed})

@app.route("/download/<jid>/<filename>")
@login_required
def download(jid, filename):
    j = get_job(jid)
    if not j: abort(404)
    target = next((Path(o["path"]) for o in j["outputs"] if o["filename"] == filename), None)
    if not target or not target.exists(): abort(404)
    return send_file(str(target), as_attachment=True, download_name=filename)

if __name__ == "__main__":
    print(f"\n  PILLAR | Auth: {'MSAL Azure AD' if CLIENT_ID else 'BYPASS locale'} | http://127.0.0.1:5000\n")
    app.run(debug=True, host="0.0.0.0", port=5000)
