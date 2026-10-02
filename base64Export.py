import base64
import binascii
import hashlib
import json
import logging
import os
import sys
import tempfile
import time
from typing import Any, Dict, Optional, Tuple

from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("sheet_sync")


def load_config(path: str) -> Dict[str, Any]:
    with open(path, "r", encoding="utf-8") as f:
        cfg = json.load(f)

    required = ["spreadsheet_id", "sheet_name", "fields"]
    missing = [k for k in required if k not in cfg]
    if missing:
        raise ValueError(f"config.json is missing required keys: {missing}")

    cfg.setdefault("credentials_file", "service_account.json")
    cfg.setdefault("poll_interval_seconds", 15)
    cfg.setdefault("output_dir", "output")
    cfg.setdefault("text_output_file", "data.json")
    cfg.setdefault("image_output_subdir", "images")
    return cfg


def build_sheets_service(credentials_file: str):
    if not os.path.exists(credentials_file):
        log.error(
            "Credentials file '%s' not found. Create a service account key "
            "in Google Cloud Console and share the sheet with its email.",
            credentials_file,
        )
        sys.exit(1)
    creds = service_account.Credentials.from_service_account_file(
        credentials_file, scopes=SCOPES
    )
    return build("sheets", "v4", credentials=creds, cache_discovery=False)


def atomic_write_json(path: str, data: Dict[str, Any]) -> None:
    dir_name = os.path.dirname(os.path.abspath(path)) or "."
    fd, tmp_path = tempfile.mkstemp(dir=dir_name, prefix=".tmp_", suffix=".json")
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        os.replace(tmp_path, path)
    except Exception:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise


def atomic_write_bytes(path: str, data: bytes) -> None:
    dir_name = os.path.dirname(os.path.abspath(path)) or "."
    fd, tmp_path = tempfile.mkstemp(dir=dir_name, prefix=".tmp_", suffix=".png")
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(data)
        os.replace(tmp_path, path)
    except Exception:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)
        raise


def extract_image_base64(cell_data: Dict[str, Any]) -> Optional[str]:
    """
    Pull the raw base64 text out of a cell.

    Expects the cell's plain text value to be the base64 string itself.
    Tolerates an optional data-URL prefix (e.g. "data:image/png;base64,")
    and strips surrounding whitespace/newlines, since long strings pasted
    into Sheets can sometimes pick up formatting artifacts.
    """
    formatted = cell_data.get("formattedValue", "")
    if not formatted:
        return None

    text = formatted.strip()
    if not text:
        return None

    if text.startswith("data:") and "," in text:
        text = text.split(",", 1)[1]

    return text


def decode_base64_image(b64_text: str) -> Optional[bytes]:
    """Decode a base64 string to raw bytes, returning None on failure."""
    try:
        # Strip any stray whitespace/newlines that may have crept in.
        cleaned = "".join(b64_text.split())
        return base64.b64decode(cleaned, validate=False)
    except (binascii.Error, ValueError) as e:
        log.warning("Could not decode base64 image data: %s", e)
        return None


def decode_if_changed(
    b64_text: str, prior_state: Dict[str, Any]
) -> Tuple[Optional[bytes], Dict[str, Any], bool, bool]:
    """
    Decode a cell's base64 text, but only report it as "changed" (and
    therefore worth writing to disk) if the decoded bytes differ from
    last cycle's. No network call is involved -- the base64 text already
    came back as part of the same Sheets API read used for every other
    field, so this is purely local decode + hash comparison.

    Returns (new_bytes_or_None, updated_state, changed_bool, decode_failed_bool).
    """
    content = decode_base64_image(b64_text)
    if content is None:
        return None, prior_state, False, True

    content_hash = hashlib.sha256(content).hexdigest()
    new_state = {"content_hash": content_hash}

    if content_hash == prior_state.get("content_hash"):
        return None, new_state, False, False

    return content, new_state, True, False


def fetch_cell_grid(service, spreadsheet_id: str, sheet_name: str, cell_refs):
    """Fetch grid data for a set of A1 cell refs in one API call."""
    ranges = [f"{sheet_name}!{cell}" for cell in cell_refs]
    resp = (
        service.spreadsheets()
        .get(
            spreadsheetId=spreadsheet_id,
            ranges=ranges,
            fields=(
                "sheets(data(rowData(values("
                "formattedValue,userEnteredValue,effectiveValue,hyperlink"
                "))))"
            ),
        )
        .execute()
    )

    # Response has one "sheets[].data[]" entry per requested range, in order.
    results = {}
    sheets_data = resp.get("sheets", [{}])[0].get("data", [])
    for cell_ref, data_block in zip(cell_refs, sheets_data):
        row_data = data_block.get("rowData", [])
        if row_data and row_data[0].get("values"):
            results[cell_ref] = row_data[0]["values"][0]
        else:
            results[cell_ref] = {}
    return results


def sync_once(
    service,
    cfg: Dict[str, Any],
    text_state: Dict[str, Any],
    image_state: Dict[str, Any],
    cycle_count: int,
) -> None:
    spreadsheet_id = cfg["spreadsheet_id"]
    sheet_name = cfg["sheet_name"]
    fields = cfg["fields"]

    image_every_n = max(1, int(cfg.get("image_update_every_n_cycles", 1)))
    # Update images on the first cycle (so files exist right away), then
    # every Nth cycle after that.
    should_update_images = (cycle_count % image_every_n == 0)

    cell_refs = [f["cell"] for f in fields]
    cell_data_map = fetch_cell_grid(service, spreadsheet_id, sheet_name, cell_refs)

    output_dir = cfg["output_dir"]
    image_dir = os.path.join(output_dir, cfg["image_output_subdir"])
    os.makedirs(image_dir, exist_ok=True)

    for field in fields:
        cell = field["cell"]
        name = field["name"]
        ftype = field.get("type", "text")
        cell_data = cell_data_map.get(cell, {})

        if ftype == "text":
            value = cell_data.get("formattedValue", "")
            text_state[name] = value

        elif ftype == "image":
            if not should_update_images:
                continue

            b64_text = extract_image_base64(cell_data)
            if not b64_text:
                log.warning(
                    "No base64 image data found in cell %s (field '%s'). "
                    "Skipping (previous file, if any, is left untouched).",
                    cell,
                    name,
                )
                continue

            prior = image_state.get(name, {})
            img_bytes, new_state, changed, decode_failed = decode_if_changed(
                b64_text, prior
            )
            image_state[name] = new_state

            if not changed:
                if decode_failed:
                    log.warning(
                        "Cell %s (field '%s') had base64 text but it failed "
                        "to decode; leaving previous file untouched.",
                        cell,
                        name,
                    )
                else:
                    log.info("Image unchanged, skipped write: %s", name)
                continue

            out_path = os.path.join(image_dir, f"{name}.png")
            atomic_write_bytes(out_path, img_bytes)
            log.info("Image changed, updated: %s", out_path)

        else:
            log.warning("Unknown field type '%s' for cell %s, skipping", ftype, cell)

    text_path = os.path.join(output_dir, cfg["text_output_file"])
    atomic_write_json(text_path, text_state)
    log.info("Updated text JSON: %s", text_path)


def main():
    config_path = sys.argv[1] if len(sys.argv) > 1 else "config.json"
    cfg = load_config(config_path)
    service = build_sheets_service(cfg["credentials_file"])

    os.makedirs(cfg["output_dir"], exist_ok=True)
    text_state: Dict[str, Any] = {}
    cycle_count = 0

    # Persisted so that restarting the script doesn't force a rewrite of
    # every image file -- we remember each field's last content hash on disk.
    image_state_path = os.path.join(cfg["output_dir"], ".image_state.json")
    image_state: Dict[str, Any] = {}
    if os.path.exists(image_state_path):
        try:
            with open(image_state_path, "r", encoding="utf-8") as f:
                image_state = json.load(f)
        except (json.JSONDecodeError, OSError) as e:
            log.warning("Could not read %s, starting fresh: %s", image_state_path, e)

    image_every_n = max(1, int(cfg.get("image_update_every_n_cycles", 1)))
    log.info(
        "Starting sync loop for spreadsheet '%s' every %ss "
        "(images decoded from base64 every %d cycle(s) = ~%ds, "
        "only written to disk when changed) "
        "(Ctrl+C to stop)",
        cfg["spreadsheet_id"],
        cfg["poll_interval_seconds"],
        image_every_n,
        image_every_n * cfg["poll_interval_seconds"],
    )

    while True:
        start = time.monotonic()
        try:
            sync_once(service, cfg, text_state, image_state, cycle_count)
            atomic_write_json(image_state_path, image_state)
        except HttpError as e:
            log.error("Google Sheets API error: %s", e)
        except Exception as e:
            log.exception("Unexpected error during sync: %s", e)

        cycle_count += 1

        elapsed = time.monotonic() - start
        sleep_for = max(0.0, cfg["poll_interval_seconds"] - elapsed)
        try:
            time.sleep(sleep_for)
        except KeyboardInterrupt:
            log.info("Stopped by user.")
            break


if __name__ == "__main__":
    main()