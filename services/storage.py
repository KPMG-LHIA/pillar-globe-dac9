"""
storage.py - Astrazione storage per PILLAR

Obiettivo: rimuovere la dipendenza dal filesystem locale di App Service F1
(effimero: i file spariscono ad ogni riavvio/idle-recycle), senza rompere
lo sviluppo locale.

Modalità:
- Se la env var AZURE_STORAGE_CONNECTION_STRING è presente → usa Azure Blob
  Storage (container "pillar-files" per i file, "pillar-jobs" per lo stato
  dei job, come blob JSON).
- Se assente → fallback trasparente su filesystem locale (comportamento
  identico alla versione precedente). Questo permette di continuare a
  sviluppare/testare in locale senza un account Azure Storage.

Uso:
    from services.storage import storage
    storage.save_file(local_path, "uploads/nome.xml")
    storage.download_to(  "uploads/nome.xml", dest_local_path)
    storage.save_job(jid, job_dict)
    job = storage.load_job(jid)

Hardening (2026-09, fix CodeQL alert #2/#3 "Uncontrolled data used in path
expression"): jid e blob_name arrivano da input utente (parametri di route
Flask: /status/<jid>, /download/<jid>/<filename>) e finiscono nella
costruzione di path filesystem/blob. Senza validazione, un jid tipo
"../../etc/passwd" o un blob_name con ".." permetterebbe path traversal
(lettura/scrittura fuori dalle directory previste). Ogni ingresso pubblico
valida esplicitamente prima di toccare il filesystem o l'SDK Blob.
"""
import json
import logging
import os
import re
import shutil
import uuid
from pathlib import Path

logger = logging.getLogger("pillar.storage")

CONN_STR = os.getenv("AZURE_STORAGE_CONNECTION_STRING", "")
FILES_CONTAINER = os.getenv("PILLAR_FILES_CONTAINER", "pillar-files")
JOBS_CONTAINER = os.getenv("PILLAR_JOBS_CONTAINER", "pillar-jobs")

BASE_DIR = Path(__file__).parent.parent
LOCAL_FILES_DIR = BASE_DIR / "uploads_local"
LOCAL_JOBS_DIR = BASE_DIR / "jobs_local"


def _validate_jid(jid: str) -> str:
    """jid deve essere uno UUID valido (è così che viene sempre generato in
    app.py:new_job). Qualsiasi altro valore è rifiutato prima di costruire
    un path — elimina il path traversal alla radice."""
    try:
        return str(uuid.UUID(str(jid)))
    except (ValueError, AttributeError, TypeError):
        raise ValueError(f"jid non valido: {jid!r}")


# blob_name legittimi hanno sempre la forma "uploads/<uuid>/<file>.xml" o
# "results/<uuid>/<stem>/<file>" — niente "..", niente path assoluti,
# niente caratteri che permettano di uscire dalla directory di base.
_SAFE_BLOB_RE = re.compile(r"^[A-Za-z0-9_.\-/]+$")


def _validate_blob_name(blob_name: str) -> str:
    blob_name = str(blob_name)
    if (
        not blob_name
        or ".." in blob_name
        or blob_name.startswith("/")
        or not _SAFE_BLOB_RE.match(blob_name)
    ):
        raise ValueError(f"blob_name non valido: {blob_name!r}")
    return blob_name


class _LocalBackend:
    """Fallback filesystem locale — usato se non è configurato Azure Storage."""

    def __init__(self):
        LOCAL_FILES_DIR.mkdir(exist_ok=True)
        LOCAL_JOBS_DIR.mkdir(exist_ok=True)
        logger.warning(
            "AZURE_STORAGE_CONNECTION_STRING non impostata: uso filesystem "
            "locale (NON persistente su App Service F1). Solo per sviluppo."
        )

    # -- file binari (upload xml, output xlsx/msg/html) ----------------------
    def save_file(self, local_src: Path, blob_name: str) -> str:
        blob_name = _validate_blob_name(blob_name)
        dest = LOCAL_FILES_DIR / blob_name
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(local_src, dest)
        return blob_name

    def download_to(self, blob_name: str, local_dest: Path) -> Path:
        blob_name = _validate_blob_name(blob_name)
        src = LOCAL_FILES_DIR / blob_name
        local_dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(src, local_dest)
        return local_dest

    def exists(self, blob_name: str) -> bool:
        blob_name = _validate_blob_name(blob_name)
        return (LOCAL_FILES_DIR / blob_name).exists()

    def open_read(self, blob_name: str):
        blob_name = _validate_blob_name(blob_name)
        return open(LOCAL_FILES_DIR / blob_name, "rb")

    # -- job state (JSON) ------------------------------------------------------
    def save_job(self, jid: str, data: dict):
        jid = _validate_jid(jid)
        path = LOCAL_JOBS_DIR / f"{jid}.json"
        path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")

    def load_job(self, jid: str):
        jid = _validate_jid(jid)
        path = LOCAL_JOBS_DIR / f"{jid}.json"
        if not path.exists():
            return None
        return json.loads(path.read_text(encoding="utf-8"))

    def delete_job(self, jid: str):
        jid = _validate_jid(jid)
        path = LOCAL_JOBS_DIR / f"{jid}.json"
        if path.exists():
            path.unlink()


class _BlobBackend:
    """Backend Azure Blob Storage — usato quando è configurata la connection string."""

    def __init__(self, conn_str: str):
        from azure.storage.blob import BlobServiceClient
        self._svc = BlobServiceClient.from_connection_string(conn_str)
        for container in (FILES_CONTAINER, JOBS_CONTAINER):
            try:
                self._svc.create_container(container)
            except Exception:
                pass  # esiste già
        self._files = self._svc.get_container_client(FILES_CONTAINER)
        self._jobs = self._svc.get_container_client(JOBS_CONTAINER)
        logger.info("Storage backend: Azure Blob Storage (container '%s', '%s')",
                    FILES_CONTAINER, JOBS_CONTAINER)

    def save_file(self, local_src: Path, blob_name: str) -> str:
        blob_name = _validate_blob_name(blob_name)
        with open(local_src, "rb") as f:
            self._files.upload_blob(blob_name, f, overwrite=True)
        return blob_name

    def download_to(self, blob_name: str, local_dest: Path) -> Path:
        blob_name = _validate_blob_name(blob_name)
        local_dest.parent.mkdir(parents=True, exist_ok=True)
        stream = self._files.download_blob(blob_name)
        with open(local_dest, "wb") as f:
            f.write(stream.readall())
        return local_dest

    def exists(self, blob_name: str) -> bool:
        blob_name = _validate_blob_name(blob_name)
        return self._files.get_blob_client(blob_name).exists()

    def open_read(self, blob_name: str):
        blob_name = _validate_blob_name(blob_name)
        import io
        stream = self._files.download_blob(blob_name)
        return io.BytesIO(stream.readall())

    def save_job(self, jid: str, data: dict):
        jid = _validate_jid(jid)
        payload = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self._jobs.upload_blob(f"{jid}.json", payload, overwrite=True)

    def load_job(self, jid: str):
        jid = _validate_jid(jid)
        client = self._jobs.get_blob_client(f"{jid}.json")
        if not client.exists():
            return None
        return json.loads(client.download_blob().readall().decode("utf-8"))

    def delete_job(self, jid: str):
        jid = _validate_jid(jid)
        client = self._jobs.get_blob_client(f"{jid}.json")
        if client.exists():
            client.delete_blob()


def _init_backend():
    if CONN_STR:
        try:
            return _BlobBackend(CONN_STR)
        except Exception as e:
            logger.error("Impossibile inizializzare Azure Blob Storage (%s). "
                         "Fallback su filesystem locale.", e)
            return _LocalBackend()
    return _LocalBackend()


storage = _init_backend()