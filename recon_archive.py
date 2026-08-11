"""Store generated reconciliation workbooks in Google Drive."""

from __future__ import annotations

import hashlib
import io
import json
import os
import re
import uuid
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Mapping


INDEX_NAME = "recon_archive_index.json"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
JSON_MIME = "application/json"
DRIVE_FOLDER_MIME = "application/vnd.google-apps.folder"
UPLOAD_CHUNK_SIZE = 5 * 1024 * 1024
UPLOAD_RETRIES = 3

CATEGORY_FOLDER_NAMES = {
    "Collection": "Collection Recon",
    "Disbursement": "Disbursement Recon",
    "Filtering": "Filtering",
    "Summary": "Summary",
}


@dataclass(frozen=True)
class DriveConfig:
    folder_id: str
    service_account_json: str | Mapping[str, Any] | None = None
    service_account_file: str | None = None
    delegated_user: str | None = None


@dataclass(frozen=True)
class ArchiveWriteResult:
    record: dict[str, Any]
    drive_saved: bool
    drive_error: str | None = None


@dataclass(frozen=True)
class DrivePruneResult:
    removed: int
    errors: list[str]


@dataclass(frozen=True)
class DriveRebuildResult:
    recovered: int
    skipped: int
    errors: list[str]


@dataclass(frozen=True)
class ArchiveDeleteResult:
    deleted: int
    errors: list[str]


def default_drive_config() -> DriveConfig | None:
    folder_id = os.environ.get("GOOGLE_DRIVE_FOLDER_ID", "").strip()
    service_account_json = os.environ.get("GOOGLE_DRIVE_SERVICE_ACCOUNT_JSON", "").strip()
    service_account_file = (
        os.environ.get("GOOGLE_DRIVE_SERVICE_ACCOUNT_FILE", "").strip()
        or os.environ.get("GOOGLE_APPLICATION_CREDENTIALS", "").strip()
    )
    delegated_user = (
        os.environ.get("GOOGLE_DRIVE_DELEGATED_USER", "").strip()
        or os.environ.get("GOOGLE_WORKSPACE_DELEGATED_USER", "").strip()
    )
    if not folder_id:
        return None
    return DriveConfig(
        folder_id=folder_id,
        service_account_json=service_account_json or None,
        service_account_file=service_account_file or None,
        delegated_user=delegated_user or None,
    )


def archive_generated_file(
    *,
    content: bytes,
    file_name: str,
    workflow: str,
    category: str,
    month: str,
    archive_scope: str = "",
    drive_config: DriveConfig | None = None,
    client: GoogleDriveArchiveClient | None = None,
) -> ArchiveWriteResult:
    """Persist a generated workbook and append it to the archive index."""
    now = datetime.now().astimezone().replace(microsecond=0).isoformat()
    record_id = uuid.uuid4().hex
    clean_name = _safe_file_name(file_name)
    stored_name = f"{month}_{_safe_file_name(workflow)}_{record_id[:8]}_{clean_name}"
    content_hash = hashlib.sha256(content).hexdigest()
    record = {
        "id": record_id,
        "month": month,
        "category": category,
        "workflow": workflow,
        "archive_scope": archive_scope,
        "file_name": clean_name,
        "archive_key": _archive_record_key({
            "month": month,
            "category": category,
            "workflow": workflow,
            "file_name": clean_name,
        }),
        "version_status": "current",
        "superseded_by": "",
        "superseded_at": "",
        "content_hash": content_hash,
        "stored_name": stored_name,
        "local_path": "",
        "mime_type": XLSX_MIME,
        "size_bytes": len(content),
        "created_at": now,
        "drive_file_id": "",
        "drive_web_view_link": "",
        "drive_saved": False,
        "drive_error": "",
    }

    drive_error: str | None = None
    drive_saved = False
    uploaded_file_id = ""

    if client is not None or drive_config:
        try:
            active_client = client if client is not None else GoogleDriveArchiveClient(drive_config)
            existing_records = normalize_archive_versions(active_client.load_index())
            existing_current = _find_current_record_with_hash(
                existing_records,
                _archive_record_key(record),
                content_hash,
            )
            if existing_current:
                return ArchiveWriteResult(record=existing_current, drive_saved=True)

            drive_file = active_client.upload_bytes(stored_name, content, XLSX_MIME, month=month, category=category)
            uploaded_file_id = str(drive_file.get("id", ""))
            record["drive_file_id"] = uploaded_file_id
            record["drive_web_view_link"] = drive_file.get("webViewLink", "")
            if not uploaded_file_id:
                raise RuntimeError("Google Drive upload did not return a file ID.")

            indexed_record = dict(record)
            indexed_record["drive_saved"] = True
            records = _records_with_new_current(existing_records, indexed_record, now)
            active_client.save_index(_drive_index_records(records))
            record.update(indexed_record)
            drive_saved = True

        except Exception as exc:  # Google auth/API is optional at runtime.
            drive_error = _format_drive_error(exc, drive_config)
            if uploaded_file_id:
                try:
                    active_client.service.files().delete(
                        fileId=uploaded_file_id,
                        supportsAllDrives=True,
                    ).execute()
                    record["drive_file_id"] = ""
                    record["drive_web_view_link"] = ""
                except Exception as cleanup_exc:
                    drive_error = f"{drive_error}; uploaded file cleanup failed: {cleanup_exc}"
            record["drive_error"] = drive_error
    else:
        drive_error = "Google Drive is not configured; the workbook was not archived."
        record["drive_error"] = drive_error

    return ArchiveWriteResult(record=record, drive_saved=drive_saved, drive_error=drive_error)


def list_archive_records(drive_config: DriveConfig | None = None) -> tuple[list[dict[str, Any]], str | None]:
    if not drive_config:
        return [], "Google Drive is not configured."

    try:
        client = GoogleDriveArchiveClient(drive_config)
        records = normalize_archive_versions(client.load_index())
        return sorted(records, key=_record_sort_key, reverse=True), None
    except Exception as exc:
        return [], _format_drive_error(exc, drive_config)


def read_archive_file(record: Mapping[str, Any], drive_config: DriveConfig | None = None) -> bytes:
    drive_file_id = str(record.get("drive_file_id", "")).strip()
    if drive_config and drive_file_id:
        return GoogleDriveArchiveClient(drive_config).download_bytes(drive_file_id)

    raise FileNotFoundError(f"Archived file is not available in Google Drive: {record.get('file_name', 'unknown file')}")


def prune_missing_drive_records(drive_config: DriveConfig) -> DrivePruneResult:
    """Drop archive records whose Drive file was deleted outside the app."""
    client = GoogleDriveArchiveClient(drive_config)
    records = client.load_index()
    kept: list[dict[str, Any]] = []
    removed = 0
    errors: list[str] = []

    for record in records:
        drive_file_id = str(record.get("drive_file_id") or "").strip()
        if not drive_file_id:
            kept.append(record)
            continue
        try:
            file_meta = client.service.files().get(
                fileId=drive_file_id, fields="id,trashed", supportsAllDrives=True
            ).execute()
            if file_meta.get("trashed"):
                removed += 1
            else:
                kept.append(record)
        except Exception as exc:
            if _is_not_found(exc):
                removed += 1
            else:
                errors.append(f"{record.get('file_name', 'Unknown file')}: {exc}")
                kept.append(record)

    if removed:
        try:
            client.save_index(_drive_index_records(normalize_archive_versions(kept)))
        except Exception as exc:
            errors.append(f"Drive index update failed: {exc}")

    return DrivePruneResult(removed=removed, errors=errors)


def rebuild_drive_index(
    drive_config: DriveConfig,
    archive_specs: list[tuple[str, str, str]],
) -> DriveRebuildResult:
    """Recreate the archive index from workbooks already stored in Drive."""
    client = GoogleDriveArchiveClient(drive_config)
    client.ping()
    records: list[dict[str, Any]] = []
    skipped = 0
    errors: list[str] = []
    category_by_folder = {
        folder_name: category for category, folder_name in CATEGORY_FOLDER_NAMES.items()
    }
    # Categories such as Ledger intentionally use their category name as the
    # folder name and therefore are not present in CATEGORY_FOLDER_NAMES.
    for category, _workflow, _file_name in archive_specs:
        category_by_folder.setdefault(category, category)

    try:
        month_folders = client.list_children(client.config.folder_id, folders_only=True)
    except Exception as exc:
        return DriveRebuildResult(0, 0, [f"Could not list archive folders: {exc}"])

    for month_folder in month_folders:
        month = str(month_folder.get("name", "")).strip()
        if not re.fullmatch(r"\d{4}-(?:0[1-9]|1[0-2])", month):
            continue
        try:
            category_folders = client.list_children(str(month_folder["id"]), folders_only=True)
        except Exception as exc:
            errors.append(f"{month}: {exc}")
            continue
        for category_folder in category_folders:
            folder_name = str(category_folder.get("name", "")).strip()
            category = category_by_folder.get(folder_name)
            if not category:
                continue
            try:
                files = client.list_children(str(category_folder["id"]), folders_only=False)
            except Exception as exc:
                errors.append(f"{month}/{folder_name}: {exc}")
                continue
            for file_meta in files:
                if file_meta.get("mimeType") == DRIVE_FOLDER_MIME:
                    continue
                recovered = _record_from_drive_file(file_meta, month, category, archive_specs)
                if recovered is None:
                    skipped += 1
                    continue
                records.append(recovered)

    if not records:
        return DriveRebuildResult(
            0,
            skipped,
            errors + ["No recognizable archived workbooks were found; the index was not changed."],
        )
    try:
        normalized = normalize_archive_versions(records)
        client.save_index(_drive_index_records(normalized))
    except Exception as exc:
        errors.append(f"Could not save the rebuilt archive index: {exc}")
        return DriveRebuildResult(0, skipped, errors)
    return DriveRebuildResult(len(records), skipped, errors)


def delete_archive_record(record: Mapping[str, Any], drive_config: DriveConfig | None = None) -> str | None:
    """Delete an archived file from Drive (if present) and the archive index. Returns an error message, if any."""
    result = delete_archive_records([record], drive_config)
    if result.errors:
        return "; ".join(result.errors)
    return None


def delete_archive_records(
    records_to_delete: list[Mapping[str, Any]],
    drive_config: DriveConfig | None = None,
) -> ArchiveDeleteResult:
    """Delete multiple archive records and their Drive files."""
    if not drive_config:
        return ArchiveDeleteResult(0, ["Google Drive is not configured."])

    client = GoogleDriveArchiveClient(drive_config)
    deleted_ids: set[str] = set()
    errors: list[str] = []

    for record in records_to_delete:
        record_id = str(record.get("id", "")).strip()
        drive_file_id = str(record.get("drive_file_id") or "").strip()
        if not record_id:
            continue
        if not drive_file_id:
            deleted_ids.add(record_id)
            continue
        try:
            client.service.files().delete(fileId=drive_file_id, supportsAllDrives=True).execute()
        except Exception as exc:
            if not _is_not_found(exc):
                errors.append(f"{record.get('file_name', 'Unknown file')}: {exc}")
                continue
        deleted_ids.add(record_id)

    if deleted_ids:
        try:
            drive_records = [
                r for r in client.load_index()
                if str(r.get("id", "")).strip() not in deleted_ids
            ]
            client.save_index(_drive_index_records(normalize_archive_versions(drive_records)))
        except Exception as exc:
            errors.append(_format_drive_error(exc, drive_config))

    return ArchiveDeleteResult(len(deleted_ids), errors)


def check_drive_connection(drive_config: DriveConfig | None = None) -> tuple[bool, str]:
    if not drive_config:
        return False, "Google Drive is not configured."
    try:
        GoogleDriveArchiveClient(drive_config).ping()
    except Exception as exc:
        return False, _format_drive_error(exc, drive_config)
    return True, "Google Drive connected."


class GoogleDriveArchiveClient:
    def __init__(self, config: DriveConfig):
        if not config.folder_id:
            raise RuntimeError("GOOGLE_DRIVE_FOLDER_ID is missing.")
        self.config = config
        self.service = self._build_service()
        self._folder_cache: dict[tuple[str, str], str] = {}
        self._folder_metadata_cache: dict[str, dict[str, Any]] = {}
        self._index_file_id: str | None = None

    def ping(self) -> None:
        self._validate_archive_root()

    def upload_bytes(
        self,
        file_name: str,
        content: bytes,
        mime_type: str,
        month: str | None = None,
        category: str | None = None,
    ) -> dict[str, Any]:
        self._validate_archive_root()
        media = self._media_upload(content, mime_type)
        parent_id = self.config.folder_id
        if month:
            parent_id = self._get_or_create_subfolder(parent_id, month)
            if category:
                category_name = CATEGORY_FOLDER_NAMES.get(category, category)
                parent_id = self._get_or_create_subfolder(parent_id, category_name)
        metadata = {"name": file_name, "parents": [parent_id]}
        return self.service.files().create(
            body=metadata,
            media_body=media,
            fields="id,name,webViewLink,webContentLink",
            supportsAllDrives=True,
        ).execute(num_retries=UPLOAD_RETRIES)

    def _get_or_create_subfolder(self, parent_id: str, name: str) -> str:
        cache_key = (parent_id, name)
        if cache_key in self._folder_cache:
            return self._folder_cache[cache_key]

        query = (
            f"'{parent_id}' in parents and "
            f"name = '{name}' and mimeType = '{DRIVE_FOLDER_MIME}' and trashed = false"
        )
        result = self.service.files().list(
            q=query,
            spaces="drive",
            fields="files(id,name)",
            pageSize=1,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
        files = result.get("files", [])
        if files:
            folder_id = files[0]["id"]
        else:
            folder = self.service.files().create(
                body={"name": name, "mimeType": DRIVE_FOLDER_MIME, "parents": [parent_id]},
                fields="id",
                supportsAllDrives=True,
            ).execute()
            folder_id = folder["id"]

        self._folder_cache[cache_key] = folder_id
        return folder_id

    def download_bytes(self, file_id: str) -> bytes:
        request = self.service.files().get_media(fileId=file_id, supportsAllDrives=True)
        buffer = io.BytesIO()
        downloader = self._media_download(buffer, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        return buffer.getvalue()

    def list_children(self, parent_id: str, *, folders_only: bool | None = None) -> list[dict[str, Any]]:
        query_parts = [f"'{parent_id}' in parents", "trashed = false"]
        if folders_only is True:
            query_parts.append(f"mimeType = '{DRIVE_FOLDER_MIME}'")
        elif folders_only is False:
            query_parts.append(f"mimeType != '{DRIVE_FOLDER_MIME}'")
        files: list[dict[str, Any]] = []
        page_token: str | None = None
        while True:
            result = self.service.files().list(
                q=" and ".join(query_parts),
                spaces="drive",
                fields=(
                    "nextPageToken,files("
                    "id,name,mimeType,size,createdTime,modifiedTime,"
                    "webViewLink,md5Checksum)"
                ),
                pageSize=1000,
                pageToken=page_token,
                supportsAllDrives=True,
                includeItemsFromAllDrives=True,
            ).execute()
            files.extend(result.get("files", []))
            page_token = result.get("nextPageToken")
            if not page_token:
                return files

    def load_index(self) -> list[dict[str, Any]]:
        file_id = self._find_index_file_id()
        if not file_id:
            return []
        try:
            payload = self.download_bytes(file_id)
            data = json.loads(payload.decode("utf-8"))
        except Exception:
            return []
        if isinstance(data, dict):
            records = data.get("records", [])
        else:
            records = data
        return [dict(record) for record in records if isinstance(record, Mapping)]

    def save_index(self, records: list[dict[str, Any]]) -> None:
        payload = json.dumps(
            {
                "updated_at": datetime.now().astimezone().replace(microsecond=0).isoformat(),
                "records": records,
            },
            indent=2,
            sort_keys=True,
        ).encode("utf-8")
        media = self._media_upload(payload, JSON_MIME)
        file_id = self._find_index_file_id()
        if file_id:
            self.service.files().update(
                fileId=file_id,
                media_body=media,
                fields="id,name",
                supportsAllDrives=True,
            ).execute()
            return
        created = self.service.files().create(
            body={"name": INDEX_NAME, "parents": [self.config.folder_id]},
            media_body=media,
            fields="id,name",
            supportsAllDrives=True,
        ).execute()
        self._index_file_id = created.get("id")

    def _find_index_file_id(self) -> str | None:
        if self._index_file_id:
            return self._index_file_id

        query = (
            f"'{self.config.folder_id}' in parents and "
            f"name = '{INDEX_NAME}' and trashed = false"
        )
        result = self.service.files().list(
            q=query,
            spaces="drive",
            fields="files(id,name,modifiedTime)",
            pageSize=1,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
        ).execute()
        files = result.get("files", [])
        if files:
            self._index_file_id = files[0]["id"]
            return self._index_file_id
        return None

    def _build_service(self):
        try:
            from google.oauth2 import service_account
            from googleapiclient.discovery import build
        except ImportError as exc:
            raise RuntimeError(
                "Install google-api-python-client and google-auth to enable Google Drive storage."
            ) from exc

        info = _service_account_info(self.config)
        if info:
            credentials = service_account.Credentials.from_service_account_info(
                info,
                scopes=["https://www.googleapis.com/auth/drive"],
            )
        else:
            service_account_file = self.config.service_account_file
            if not service_account_file:
                raise RuntimeError(
                    "Provide GOOGLE_DRIVE_SERVICE_ACCOUNT_JSON or GOOGLE_DRIVE_SERVICE_ACCOUNT_FILE."
                )
            credentials = service_account.Credentials.from_service_account_file(
                service_account_file,
                scopes=["https://www.googleapis.com/auth/drive"],
            )
        if self.config.delegated_user:
            credentials = credentials.with_subject(self.config.delegated_user)
        return build("drive", "v3", credentials=credentials, cache_discovery=False)

    @staticmethod
    def _media_upload(content: bytes, mime_type: str):
        from googleapiclient.http import MediaIoBaseUpload

        # Large filtering workbooks (Nsano's combined disb+collection export can
        # run 50-80MB) are unreliable as a single non-resumable POST — chunked,
        # resumable uploads tolerate network drops by resuming instead of
        # restarting the whole upload.
        resumable = len(content) > UPLOAD_CHUNK_SIZE
        return MediaIoBaseUpload(
            io.BytesIO(content),
            mimetype=mime_type,
            resumable=resumable,
            chunksize=UPLOAD_CHUNK_SIZE if resumable else -1,
        )

    @staticmethod
    def _media_download(buffer: io.BytesIO, request):
        from googleapiclient.http import MediaIoBaseDownload

        return MediaIoBaseDownload(buffer, request)

    def _validate_archive_root(self) -> None:
        folder = self._folder_metadata(self.config.folder_id)
        if folder.get("mimeType") != DRIVE_FOLDER_MIME:
            raise RuntimeError("Configured Google Drive archive target is not a folder.")
        capabilities = folder.get("capabilities", {})
        if isinstance(capabilities, Mapping) and capabilities.get("canAddChildren") is False:
            raise RuntimeError(
                "Configured Google Drive archive folder does not allow this account to create files."
            )
        if not folder.get("driveId") and not self.config.delegated_user:
            raise RuntimeError(_service_account_quota_message())

    def _folder_metadata(self, folder_id: str) -> dict[str, Any]:
        if folder_id not in self._folder_metadata_cache:
            self._folder_metadata_cache[folder_id] = self.service.files().get(
                fileId=folder_id,
                fields="id,name,mimeType,driveId,capabilities/canAddChildren",
                supportsAllDrives=True,
            ).execute()
        return self._folder_metadata_cache[folder_id]


def _service_account_info(config: DriveConfig) -> dict[str, Any] | None:
    raw = config.service_account_json
    if isinstance(raw, Mapping):
        return dict(raw)
    if isinstance(raw, str) and raw.strip():
        text = raw.strip()
        if text.startswith("{"):
            return json.loads(text)
        path = Path(text).expanduser()
        if path.exists():
            return json.loads(path.read_text(encoding="utf-8"))
    return None


def _is_not_found(exc: Exception) -> bool:
    try:
        from googleapiclient.errors import HttpError
    except ImportError:
        return False
    return isinstance(exc, HttpError) and exc.resp.status == 404


def _format_drive_error(exc: Exception, config: DriveConfig | None = None) -> str:
    reason, message = _http_error_reason_and_message(exc)
    text = str(exc)
    if (
        reason == "storageQuotaExceeded"
        or "Service Accounts do not have storage quota" in message
        or "Service Accounts do not have storage quota" in text
    ):
        return _service_account_quota_message(config)
    return text


def _http_error_reason_and_message(exc: Exception) -> tuple[str, str]:
    try:
        from googleapiclient.errors import HttpError
    except ImportError:
        return "", ""
    if not isinstance(exc, HttpError):
        return "", ""

    try:
        raw_content = exc.content.decode("utf-8") if isinstance(exc.content, bytes) else str(exc.content)
        payload = json.loads(raw_content)
    except Exception:
        return "", ""

    error = payload.get("error", {}) if isinstance(payload, Mapping) else {}
    errors = error.get("errors", []) if isinstance(error, Mapping) else []
    reason = ""
    if errors and isinstance(errors[0], Mapping):
        reason = str(errors[0].get("reason", ""))
    message = str(error.get("message", "")) if isinstance(error, Mapping) else ""
    return reason, message


def _service_account_quota_message(config: DriveConfig | None = None) -> str:
    if config and config.delegated_user:
        delegation_hint = (
            " Domain-wide delegation is enabled in the app config, so verify that the delegated user is "
            "authorized by the Workspace admin and has access to the archive folder."
        )
    else:
        delegation_hint = (
            " Move the archive folder into a Shared Drive and add the service account as a member, "
            "or configure domain-wide delegation with GOOGLE_DRIVE_DELEGATED_USER."
        )
    return (
        "Google Drive archive upload failed because service accounts do not have personal Drive "
        "storage quota. A user-owned My Drive folder shared with the service account is not enough."
        f"{delegation_hint}"
    )


def _archive_record_key(record: Mapping[str, Any]) -> str:
    parts = [
        record.get("month", ""),
        record.get("category", ""),
        record.get("workflow", ""),
        record.get("file_name", ""),
    ]
    return "|".join(str(part or "").strip().casefold() for part in parts)


def _record_from_drive_file(
    file_meta: Mapping[str, Any],
    month: str,
    category: str,
    archive_specs: list[tuple[str, str, str]],
) -> dict[str, Any] | None:
    stored_name = str(file_meta.get("name", "")).strip()
    drive_file_id = str(file_meta.get("id", "")).strip()
    if not stored_name or not drive_file_id:
        return None

    matched_workflow = ""
    matched_file_name = ""
    for spec_category, workflow, file_name in archive_specs:
        if spec_category != category:
            continue
        safe_workflow = _safe_file_name(workflow)
        safe_name = _safe_file_name(file_name)
        pattern = rf"^{re.escape(month)}_{re.escape(safe_workflow)}_[0-9a-fA-F]{{8}}_{re.escape(safe_name)}$"
        if re.fullmatch(pattern, stored_name):
            matched_workflow = workflow
            matched_file_name = safe_name
            break
    # Keep recovery resilient to a workflow label being renamed in a later app
    # version. The stable output filename still maps it to the current label.
    if not matched_workflow:
        for spec_category, workflow, file_name in archive_specs:
            if spec_category != category:
                continue
            safe_name = _safe_file_name(file_name)
            pattern = rf"^{re.escape(month)}_.+_[0-9a-fA-F]{{8}}_{re.escape(safe_name)}$"
            if re.fullmatch(pattern, stored_name):
                matched_workflow = workflow
                matched_file_name = safe_name
                break
    if not matched_workflow:
        return None

    created_at = str(file_meta.get("createdTime") or file_meta.get("modifiedTime") or "")
    record_id = hashlib.sha256(f"drive:{drive_file_id}".encode("utf-8")).hexdigest()[:32]
    record = {
        "id": record_id,
        "month": month,
        "category": category,
        "workflow": matched_workflow,
        "archive_scope": "master",
        "file_name": matched_file_name,
        "version_status": "current",
        "superseded_by": "",
        "superseded_at": "",
        "content_hash": "",
        "stored_name": stored_name,
        "local_path": "",
        "mime_type": str(file_meta.get("mimeType") or XLSX_MIME),
        "size_bytes": int(file_meta.get("size") or 0),
        "created_at": created_at,
        "drive_file_id": drive_file_id,
        "drive_web_view_link": str(file_meta.get("webViewLink") or ""),
        "drive_saved": True,
        "drive_error": "",
    }
    record["archive_key"] = _archive_record_key(record)
    return record


def _record_is_previous(record: Mapping[str, Any]) -> bool:
    return bool(record.get("superseded_by")) or str(record.get("version_status", "")).casefold() == "previous"


def _find_current_record_with_hash(
    records: list[dict[str, Any]],
    archive_key: str,
    content_hash: str,
) -> dict[str, Any] | None:
    for record in records:
        if _archive_record_key(record) != archive_key:
            continue
        if _record_is_previous(record):
            continue
        if str(record.get("content_hash", "")).strip() != content_hash:
            continue
        if record.get("drive_saved") and record.get("drive_file_id"):
            return dict(record)
    return None


def _records_with_new_current(
    records: list[dict[str, Any]],
    new_record: dict[str, Any],
    superseded_at: str,
) -> list[dict[str, Any]]:
    archive_key = _archive_record_key(new_record)
    updated: list[dict[str, Any]] = []
    for record in records:
        current = dict(record)
        if str(current.get("id", "")).strip() != str(new_record.get("id", "")).strip() and _archive_record_key(current) == archive_key:
            current["archive_key"] = archive_key
            current["version_status"] = "previous"
            current["superseded_by"] = str(new_record.get("id", ""))
            current["superseded_at"] = superseded_at
        updated.append(current)

    current_new = dict(new_record)
    current_new["archive_key"] = archive_key
    current_new["version_status"] = "current"
    current_new["superseded_by"] = ""
    current_new["superseded_at"] = ""
    return normalize_archive_versions(_merge_records(updated, [current_new]))


def normalize_archive_versions(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    groups: dict[str, list[dict[str, Any]]] = {}
    for record in records:
        current = dict(record)
        archive_key = _archive_record_key(current)
        if not archive_key.strip("|"):
            continue
        current["archive_key"] = archive_key
        groups.setdefault(archive_key, []).append(current)

    normalized: list[dict[str, Any]] = []
    for archive_key, group in groups.items():
        ordered = sorted(group, key=_record_sort_key, reverse=True)
        if not ordered:
            continue
        current_record = dict(ordered[0])
        current_record["archive_key"] = archive_key
        current_record["version_status"] = "current"
        current_record["superseded_by"] = ""
        current_record["superseded_at"] = ""
        normalized.append(current_record)

        for previous in ordered[1:]:
            previous_record = dict(previous)
            previous_record["archive_key"] = archive_key
            previous_record["version_status"] = "previous"
            previous_record["superseded_by"] = str(current_record.get("id", ""))
            previous_record["superseded_at"] = (
                str(previous_record.get("superseded_at", "")).strip()
                or str(current_record.get("created_at", "")).strip()
            )
            normalized.append(previous_record)

    return sorted(normalized, key=_record_sort_key, reverse=True)


def _merge_records(*record_groups: list[dict[str, Any]]) -> list[dict[str, Any]]:
    merged: dict[str, dict[str, Any]] = {}
    for records in record_groups:
        for record in records:
            record_id = str(record.get("id", "")).strip()
            if not record_id:
                continue
            existing = merged.get(record_id, {})
            merged[record_id] = {**existing, **record}
    return sorted(merged.values(), key=_record_sort_key, reverse=True)


def _drive_index_records(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [
        {**record, "local_path": ""}
        for record in records
        if record.get("drive_saved") and record.get("drive_file_id")
    ]


def _record_sort_key(record: Mapping[str, Any]) -> str:
    return str(record.get("created_at", ""))


def _safe_file_name(value: str) -> str:
    text = re.sub(r"[^\w.\- ]+", "_", str(value or "file").strip())
    text = re.sub(r"\s+", " ", text).strip(" ._")
    return text or "file"
