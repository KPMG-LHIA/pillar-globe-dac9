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
"""
import json
import logging
import os
import shutil
from pathlib import Path

logger = logging.getLogger("pillar.storage")

CONN_STR = os.getenv("AZURE_STORAGE_CONNECTION_STRING", "")
FILES_CONTAINER = os.getenv("PILLAR_FILES_CONTAINER", "pillar-files")
JOBS_CONTAINER = os.getenv("PILLAR_JOBS_CONTAINER", "pillar-jobs")

BASE_DIR = Path(__file__).parent.parent
LOCAL_FILES_DIR = BASE_DIR / "uploads_local"
LOCAL_JOBS_DIR = BASE_DIR / "jobs_local"


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
        dest = LOCAL_FILES_DIR / blob_name
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(local_src, dest)
        return blob_name

    def download_to(self, blob_name: str, local_dest: Path) -> Path:
        src = LOCAL_FILES_DIR / blob_name
        local_dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(src, local_dest)
        return local_dest

    def exists(self, blob_name: str) -> bool:
        return (LOCAL_FILES_DIR / blob_name).exists()

    def open_read(self, blob_name: str):
        return open(LOCAL_FILES_DIR / blob_name, "rb")

    # -- job state (JSON) ------------------------------------------------------
    def save_job(self, jid: str, data: dict):
        path = LOCAL_JOBS_DIR / f"{jid}.json"
        path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")

    def load_job(self, jid: str):
        path = LOCAL_JOBS_DIR / f"{jid}.json"
        if not path.exists():
            return None
        return json.loads(path.read_text(encoding="utf-8"))

    def delete_job(self, jid: str):
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
        with open(local_src, "rb") as f:
            self._files.upload_blob(blob_name, f, overwrite=True)
        return blob_name

    def download_to(self, blob_name: str, local_dest: Path) -> Path:
        local_dest.parent.mkdir(parents=True, exist_ok=True)
        stream = self._files.download_blob(blob_name)
        with open(local_dest, "wb") as f:
            f.write(stream.readall())
        return local_dest

    def exists(self, blob_name: str) -> bool:
        return self._files.get_blob_client(blob_name).exists()

    def open_read(self, blob_name: str):
        import io
        stream = self._files.download_blob(blob_name)
        return io.BytesIO(stream.readall())

    def save_job(self, jid: str, data: dict):
        payload = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self._jobs.upload_blob(f"{jid}.json", payload, overwrite=True)

    def load_job(self, jid: str):
        client = self._jobs.get_blob_client(f"{jid}.json")
        if not client.exists():
            return None
        return json.loads(client.download_blob().readall().decode("utf-8"))

    def delete_job(self, jid: str):
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