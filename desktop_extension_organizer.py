import ctypes
import argparse
import json
import os
import queue
import re
import shutil
import subprocess
import sys
import threading
import tkinter as tk
import unicodedata
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog

from tkinter import ttk

try:
    import winreg
except ImportError:
    winreg = None


DESKTOP_NAMES_100_PLUS = [
    "Desktop",
    "DeskTop",
    "Masaustu",
    "Masaustu",
    "Masa\u00fcst\u00fc",
    "Escritorio",
    "Bureau",
    "Schreibtisch",
    "Scrivania",
    "Bureaublad",
    "Skrivbord",
    "Skrivebord",
    "Tyopoyta",
    "Ty\u00f6p\u00f6yt\u00e4",
    "Pulpit",
    "Plocha",
    "Pracovna plocha",
    "Asztal",
    "Birou",
    "Radna povrsina",
    "Namizje",
    "Darbvirsma",
    "Darbastalis",
    "Toolaud",
    "Skrivbordet",
    "Skrifbord",
    "Skjabor\u00f0",
    "Bwrdd gwaith",
    "Deisce",
    "Mahaigaina",
    "Escriptori",
    "\u05e9\u05d5\u05dc\u05d7\u05df \u05e2\u05d1\u05d5\u05d3\u05d4",
    "\u0633\u0637\u062d \u0627\u0644\u0645\u0643\u062a\u0628",
    "\u0645\u06cc\u0632\u06a9\u0627\u0631",
    "\u0688\u06cc\u0633\u06a9 \u0679\u0627\u067e",
    "\u0921\u0947\u0938\u094d\u0915\u091f\u0949\u092a",
    "\u09a1\u09c7\u09b8\u0995\u099f\u09aa",
    "\u0a21\u0a48\u0a38\u0a15\u0a1f\u0a3e\u0a2a",
    "\u0aaa\u0ac7\u0ab8\u0acd\u0a95\u0a9f\u0acb\u0aaa",
    "\u0921\u0947\u0938\u094d\u0915\u091f\u0949\u092a",
    "\u0ba4\u0bc6\u0bb8\u0bcd\u0b95\u0b9f\u0bbe\u0baa\u0bcd",
    "\u0c21\u0c46\u0c38\u0c4d\u0c15\u0c4d\u0c1f\u0c3e\u0c2a\u0c4d",
    "\u0ca1\u0cc6\u0cb8\u0ccd\u0c95\u0ccd\u0c9f\u0cbe\u0caa\u0ccd",
    "\u0d21\u0d46\u0d38\u0d4d\u0d15\u0d4d\u0d1f\u0d4b\u0d2a\u0d4d",
    "\u0dc3\u0dca\u0d9a\u0dca\u0dbb\u0dd3\u0db1\u0dca \u0dad\u0dbd\u0dba",
    "\u0921\u0947\u0938\u094d\u0915\u091f\u092a",
    "\u684c\u9762",
    "\u684c\u9762\u7cfb\u7d71",
    "\u30c7\u30b9\u30af\u30c8\u30c3\u30d7",
    "\ubc14\ud0d5\ud654\uba74",
    "\u0e40\u0e14\u0e2a\u0e01\u0e4c\u0e17\u0e47\u0e2d\u0e1b",
    "M\u00e0n h\u00ecnh n\u1ec1n",
    "Layar utama",
    "Desktop utama",
    "Meja Kerja",
    "Area de Trabalho",
    "\u00c1rea de Trabalho",
    "\u0420\u0430\u0431\u043e\u0447\u0438\u0439 \u0441\u0442\u043e\u043b",
    "\u0420\u0430\u0431\u043e\u0442\u0435\u043d \u043f\u043b\u043e\u0442",
    "\u0420\u043e\u0431\u043e\u0447\u0438\u0439 \u0441\u0442\u0456\u043b",
    "\u041f\u0440\u0430\u0446\u043e\u045e\u043d\u044b \u0441\u0442\u043e\u043b",
    "\u0420\u0430\u0434\u043d\u0430 \u043f\u043e\u0432\u0440\u0448\u0438\u043d\u0430",
    "\u0420\u0430\u0431\u043e\u0442\u043d\u0430 \u043f\u043e\u0432\u0440\u0448\u0438\u043d\u0430",
    "\u0531\u0577\u056d\u0561\u057f\u0561\u057d\u0565\u0572\u0561\u0576",
    "\u10e1\u10d0\u10db\u10e3\u10e8\u10d0\u10dd \u10db\u10d0\u10d2\u10d8\u10d3\u10d0",
    "I\u015f masas\u0131",
    "Ish stoli",
    "\u0416\u04b1\u043c\u044b\u0441 \u04af\u0441\u0442\u0435\u043b\u0456",
    "\u0418\u0448 \u0442\u0430\u043a\u0442\u0430",
    "\u041c\u0438\u0437\u0438 \u043a\u043e\u0440\u04e3",
    "I\u015f stoly",
    "\u0410\u0436\u043b\u044b\u043d \u0448\u0438\u0440\u044d\u044d",
    "\u1795\u17d2\u1791\u17c3\u200b\u1780\u17b6\u179a\u1784\u17b6\u179a",
    "\u0e9e\u0eb7\u0ec9\u0e99\u0e97\u0eb5\u0ec8\u0ec0\u0eae\u0eb1\u0e94\u0ea7\u0ebd\u0e81",
    "\u1012\u102e\u1038\u1005\u103a\u1000\u1010\u1031\u102c\u1037",
    "\u12f0\u1235\u12ad\u1276\u1355",
    "Desktopu",
    "Benchi ya kazi",
    "Idesikithophu",
    "\u062f\u06cc\u0633\u06a9\u0679\u0627\u067e",
    "Bord gwaith",
    "Tabili ta aiki",
    "Deskt\u00f3p",
    "Pinnalaud",
    "Skrivebordet",
    "Skrivbord",
    "Skrivbordet",
    "Arbetsyta",
    "Arbeitsfl\u00e4che",
    "\u0395\u03c0\u03b9\u03c6\u03ac\u03bd\u03b5\u03b9\u03b1 \u03b5\u03c1\u03b3\u03b1\u03c3\u03af\u03b1\u03c2",
    "\u0395\u03c0\u03b9\u03c6\u03b1\u03bd\u03b5\u03b9\u03b1\u0395\u03c1\u03b3\u03b1\u03c3\u03b9\u03b1\u03c2",
    "\u178a\u17c2\u179f\u1780\u17cb\u1790\u17bb\u1794",
    "\u0679\u06cc\u0644\u06cc\u0641\u0648\u0646 \u06c1\u0648\u0645",
    "Plocha pracovna",
    "P\u0159acovn\u00ed plocha",
    "Pracovn\u00e1 plocha",
    "\u0420\u043e\u0431\u043e\u0447\u0438\u0439\u0441\u0442\u0456\u043b",
    "\u05e9\u05d5\u05dc\u05d7\u05df\u05e2\u05d1\u05d5\u05d3\u05d4",
    "\u0645\u06cc\u0632\u06a9\u0627\u0631\u06cc",
    "\u0633\u0637\u062d\u0627\u0644\u0645\u0643\u062a\u0628",
    "\u686c\u9762",
    "\u30c7\u30b9\u30af\u30c8\u30c3\u30d7\u753b\u9762",
    "\ub370\uc2a4\ud06c\ud0d1",
    "\u0e2b\u0e19\u0e49\u0e32\u0e08\u0e2d\u0e2b\u0e25\u0e31\u0e01",
    "\u1795\u17d2\u1791\u17c3\u1795\u17d2\u1780\u17b6\u1799",
    "\u1015\u103c\u1004\u103a\u101e\u102c\u1038",
    "\u0ca1\u0cc6\u0cb8\u0ccd\u0c95\u0ccd\u0c9f\u0cbe\u0caa\u0ccd",
    "\u0d21\u0d46\u0d38\u0d4d\u0d15\u0d4d\u0d1f\u0d4b\u0d2a\u0d4d",
    "\u0baa\u0ba3\u0bbf \u0bae\u0bc7\u0b9a\u0bc8",
    "\u0c35\u0c46\u0c32\u0c4d\u0c32\u0c3e\u0c21\u0c41 \u0c2a\u0c26\u0c4d\u0c26\u0c24\u0c3f",
    "\u05de\u05e9\u05d8\u05d7 \u05e2\u05d1\u05d5\u05d3\u05d4",
    "\u0434\u0435\u0441\u043a\u0442\u043e\u043f",
    "\u043f\u043b\u043e\u0442",
    "\u684c\u9762\u76ee\u5f55",
    "\ubc14\ud0d5\ud654\uba74 \ud3f4\ub354",
    "\u0393\u03c1\u03b1\u03c6\u03b5\u03af\u03bf",
    "\u00c7al\u0131\u015fma Masas\u0131",
    "\u00c7alisma Masasi",
    "Arbeitsoberfl\u00e4che",
    "Bureau principal",
    "Escritorio principal",
    "\u00c1rea principal",
    "Main Desktop",
    "Desktop Folder",
    "Desktop Directory",
]


def _normalize_for_match(value: str) -> str:
    decomposed = unicodedata.normalize("NFKD", value)
    no_marks = "".join(ch for ch in decomposed if not unicodedata.combining(ch))
    return "".join(ch.lower() for ch in no_marks if ch.isalnum())


DESKTOP_ALIASES = {
    _normalize_for_match(name) for name in DESKTOP_NAMES_100_PLUS if name.strip()
}

UNDO_FILE_PREFIX = "desktop_organizer_undo_"
LOG_FILE_PREFIX = "desktop_organizer_log_"
TASK_NAME = "DesktopOrganizerDailyCleanup"


def _dedupe_paths(paths: list[Path]) -> list[Path]:
    unique = []
    seen = set()
    for path in paths:
        key = str(path).rstrip("\\/").casefold()
        if key not in seen:
            seen.add(key)
            unique.append(path)
    return unique


def _get_folder_path_from_shell(csidl: int) -> Path | None:
    buf = ctypes.create_unicode_buffer(260)
    result = ctypes.windll.shell32.SHGetFolderPathW(None, csidl, None, 0, buf)
    if result == 0 and buf.value:
        return Path(buf.value)
    return None


def _get_desktop_from_registry() -> Path | None:
    if winreg is None:
        return None

    registry_targets = [
        (
            winreg.HKEY_CURRENT_USER,
            r"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\User Shell Folders",
            "Desktop",
        ),
        (
            winreg.HKEY_CURRENT_USER,
            r"Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders",
            "Desktop",
        ),
    ]

    for root, key, value_name in registry_targets:
        try:
            with winreg.OpenKey(root, key) as reg_key:
                value, _ = winreg.QueryValueEx(reg_key, value_name)
                if not value:
                    continue
                expanded = os.path.expandvars(value)
                return Path(expanded)
        except OSError:
            continue

    return None


def _is_probable_desktop_dir_name(name: str) -> bool:
    normalized = _normalize_for_match(name)
    if not normalized:
        return False
    if normalized in DESKTOP_ALIASES:
        return True

    probable_keywords = [
        "desktop",
        "desk",
        "masaustu",
        "escritorio",
        "bureau",
        "schreibtisch",
        "scrivania",
        "pulpit",
        "plocha",
        "radnapovrsina",
        "rabochiistol",
        "zhuomian",
        "darbvirsma",
        "darbastalis",
        "toolaud",
        "arade",
        "skrivbord",
        "skrivebord",
    ]
    return any(keyword in normalized for keyword in probable_keywords)


def _collect_candidate_roots() -> list[Path]:
    roots: list[Path] = []
    user_home = Path.home()
    roots.append(user_home)

    env_keys = [
        "USERPROFILE",
        "PUBLIC",
        "ALLUSERSPROFILE",
        "HOMEDRIVE",
        "HOMEPATH",
        "ONEDRIVE",
        "ONEDRIVECONSUMER",
        "ONEDRIVECOMMERCIAL",
    ]

    for key in env_keys:
        value = os.environ.get(key)
        if value:
            roots.append(Path(value))

    homedrive = os.environ.get("HOMEDRIVE")
    homepath = os.environ.get("HOMEPATH")
    if homedrive and homepath:
        roots.append(Path(f"{homedrive}{homepath}"))

    if user_home.parent:
        roots.append(user_home.parent)

    cloud_root_names = [
        "OneDrive",
        "OneDrive - Personal",
        "OneDrive - Business",
        "OneDrive - Kisisel",
        "OneDrive - Sahsi",
        "Dropbox",
        "Google Drive",
        "GoogleDrive",
        "iCloudDrive",
        "YandexDisk",
        "MEGA",
        "Box",
        "Nextcloud",
        "pCloudDrive",
        "Seafile",
        "Sync",
    ]

    for root_name in cloud_root_names:
        roots.append(user_home / root_name)

    return _dedupe_paths(roots)


def get_desktop_path() -> Path:
    """Return the user's Desktop path in a locale-safe way."""
    shell_candidates = [
        _get_folder_path_from_shell(0x0010),  # CSIDL_DESKTOPDIRECTORY
        _get_folder_path_from_shell(0x0019),  # CSIDL_COMMON_DESKTOPDIRECTORY
        _get_desktop_from_registry(),
    ]
    for candidate in shell_candidates:
        if candidate and candidate.exists() and candidate.is_dir():
            return candidate

    roots = _collect_candidate_roots()
    fallback_candidates: list[Path] = []

    for root in roots:
        if _is_probable_desktop_dir_name(root.name):
            fallback_candidates.append(root)

        # Localized folder names (100+ variants) directly under known roots.
        for name in DESKTOP_NAMES_100_PLUS:
            fallback_candidates.append(root / name)

        # Catch unknown naming patterns by scanning one level deep.
        if root.exists() and root.is_dir():
            try:
                for child in root.iterdir():
                    if not child.is_dir():
                        continue
                    if _is_probable_desktop_dir_name(child.name):
                        fallback_candidates.append(child)
            except OSError:
                pass

    for candidate in _dedupe_paths(fallback_candidates):
        if candidate.exists() and candidate.is_dir():
            return candidate

    raise FileNotFoundError("Desktop klasoru bulunamadi. Geni\u015fletilmi\u015f fallback listesi ile de eri\u015filemedi.")


def _sanitize_folder_name(value: str) -> str:
    forbidden_chars = '<>:"/\\|?*'
    sanitized = value
    for char in forbidden_chars:
        sanitized = sanitized.replace(char, "_")
    sanitized = sanitized.strip().rstrip(".")
    return sanitized or "organized_files"


def _parse_excluded_extensions(raw_value: str) -> tuple[set[str], list[str]]:
    separators = [",", ";", "\n", "\t", " "]
    normalized = raw_value
    for separator in separators:
        normalized = normalized.replace(separator, ",")

    tokens = [token.strip().lower() for token in normalized.split(",") if token.strip()]
    excluded_extensions: set[str] = set()
    normalized_without_dot: list[str] = []

    for token in tokens:
        if token.startswith("."):
            excluded_extensions.add(token)
        else:
            normalized_without_dot.append(token)
            excluded_extensions.add(f".{token}")

    return excluded_extensions, normalized_without_dot


def _build_unique_destination(path: Path) -> Path:
    if not path.exists():
        return path

    stem = path.stem
    suffix = path.suffix
    counter = 1
    while True:
        candidate = path.with_name(f"{stem}_{counter}{suffix}")
        if not candidate.exists():
            return candidate
        counter += 1


def _folder_token(path: Path) -> str:
    token = "".join(ch if ch.isalnum() else "_" for ch in path.name)
    token = token.strip("_")
    return token or "folder"


def _undo_state_path(target_path: Path, stamp: str) -> Path:
    return target_path / f"{UNDO_FILE_PREFIX}{stamp}_{_folder_token(target_path)}.json"


def _latest_undo_file_for_folder(target_path: Path) -> Path | None:
    if not target_path.exists() or not target_path.is_dir():
        return None

    candidates = sorted(
        target_path.glob(f"{UNDO_FILE_PREFIX}*.json"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return candidates[0] if candidates else None


def _write_run_artifacts(
    target_path: Path,
    template: str,
    excluded_extensions: set[str],
    no_extension_to_shortcuts: bool,
    moved_count: int,
    skipped_count: int,
    total_count: int,
    operations: list[dict[str, str]],
) -> tuple[Path, Path]:
    timestamp = datetime.now()
    stamp = timestamp.strftime("%Y%m%d_%H%M%S")
    log_path = target_path / f"{LOG_FILE_PREFIX}{stamp}.txt"
    state_path = _undo_state_path(target_path, stamp)

    with log_path.open("w", encoding="utf-8") as log_file:
        log_file.write("Desktop Organizer Run Log\n")
        log_file.write(f"Timestamp: {timestamp.isoformat()}\n")
        log_file.write(f"Target folder: {target_path}\n")
        log_file.write(f"Template: {template}\n")
        log_file.write(
            "Excluded extensions: "
            + (", ".join(sorted(excluded_extensions)) if excluded_extensions else "(none)")
            + "\n"
        )
        log_file.write(f"No-extension to unknowns: {no_extension_to_shortcuts}\n")
        log_file.write(f"Total files scanned: {total_count}\n")
        log_file.write(f"Moved: {moved_count}\n")
        log_file.write(f"Skipped: {skipped_count}\n\n")
        log_file.write("Moved files:\n")
        for op in operations:
            log_file.write(f"- {op['src']} -> {op['dst']}\n")

    state_payload = {
        "timestamp": timestamp.isoformat(),
        "timestamp_label": stamp,
        "target_folder": str(target_path),
        "target_folder_name": target_path.name,
        "template": template,
        "excluded_extensions": sorted(excluded_extensions),
        "no_extension_to_unknowns": no_extension_to_shortcuts,
        "log_file": str(log_path),
        "operations": operations,
    }
    with state_path.open("w", encoding="utf-8") as state_file:
        json.dump(state_payload, state_file, ensure_ascii=False, indent=2)

    return log_path, state_path


def can_undo_for_path(target_path: Path) -> bool:
    return _latest_undo_file_for_folder(target_path) is not None


def undo_last_run(state_path: Path, progress_callback=None) -> tuple[int, int, int, str]:
    if not state_path.exists():
        raise FileNotFoundError("No undo state found for this folder.")

    with state_path.open("r", encoding="utf-8") as state_file:
        payload = json.load(state_file)

    operations = payload.get("operations", [])
    total = len(operations)
    undone = 0
    skipped = 0

    for index, op in enumerate(reversed(operations), start=1):
        src = Path(op.get("src", ""))
        dst = Path(op.get("dst", ""))

        if not dst.exists() or not dst.is_file():
            skipped += 1
            if progress_callback:
                progress_callback(index, total)
            continue

        restore_path = _build_unique_destination(src)
        try:
            restore_path.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(dst), str(restore_path))
            undone += 1
        except Exception:
            skipped += 1

        if progress_callback:
            progress_callback(index, total)

    target_folder = str(payload.get("target_folder", ""))
    state_path.unlink(missing_ok=True)
    return undone, skipped, total, target_folder


def _is_valid_hhmm(value: str) -> bool:
    if not re.match(r"^\d{2}:\d{2}$", value):
        return False
    hour = int(value[:2])
    minute = int(value[3:])
    return 0 <= hour <= 23 and 0 <= minute <= 59


def _parse_iso_date(value: str) -> datetime | None:
    try:
        return datetime.strptime(value, "%Y-%m-%d")
    except ValueError:
        return None


def _is_positive_int(value: str) -> bool:
    return value.isdigit() and int(value) > 0


def _normalize_weekday_token(token: str) -> str | None:
    normalized = _normalize_for_match(token)
    mapping = {
        "mon": "MON",
        "monday": "MON",
        "pzt": "MON",
        "pazartesi": "MON",
        "tue": "TUE",
        "tues": "TUE",
        "tuesday": "TUE",
        "sal": "TUE",
        "sali": "TUE",
        "wed": "WED",
        "wednesday": "WED",
        "car": "WED",
        "carsamba": "WED",
        "thu": "THU",
        "thur": "THU",
        "thursday": "THU",
        "per": "THU",
        "persembe": "THU",
        "fri": "FRI",
        "friday": "FRI",
        "cum": "FRI",
        "cuma": "FRI",
        "sat": "SAT",
        "saturday": "SAT",
        "cts": "SAT",
        "cumartesi": "SAT",
        "sun": "SUN",
        "sunday": "SUN",
        "paz": "SUN",
        "pazar": "SUN",
    }
    return mapping.get(normalized)


def _parse_weekdays(raw_value: str) -> list[str] | None:
    separators = [",", ";", " ", "\t", "\n"]
    normalized = raw_value
    for separator in separators:
        normalized = normalized.replace(separator, ",")

    tokens = [token.strip() for token in normalized.split(",") if token.strip()]
    if not tokens:
        return None

    result: list[str] = []
    seen: set[str] = set()
    for token in tokens:
        day = _normalize_weekday_token(token)
        if day is None:
            return None
        if day not in seen:
            seen.add(day)
            result.append(day)
    return result or None


def _quote_for_task(arg: str) -> str:
    escaped = arg.replace('"', '""')
    return f'"{escaped}"'


def _build_task_run_command(target_path: Path | None) -> str:
    args: list[str] = []
    if getattr(sys, "frozen", False):
        args.append(str(Path(sys.executable)))
    else:
        args.append(str(Path(sys.executable)))
        args.append(str(Path(__file__).resolve()))

    args.append("--auto-run")
    if target_path is not None:
        args.extend(["--path", str(target_path)])

    return " ".join(_quote_for_task(arg) for arg in args)


def create_or_update_task(schedule: dict[str, str], target_path: Path | None) -> tuple[bool, str]:
    schedule_mode = schedule.get("mode", "")
    run_command = _build_task_run_command(target_path)
    cmd = [
        "schtasks",
        "/Create",
        "/TN",
        TASK_NAME,
        "/TR",
        run_command,
        "/F",
    ]

    target_label = str(target_path) if target_path else "Desktop"

    if schedule_mode == "once":
        schedule_time = schedule.get("time", "")
        if not _is_valid_hhmm(schedule_time):
            return False, "Invalid time format. Use HH:MM."
        schedule_date = schedule.get("date", "")
        parsed_date = _parse_iso_date(schedule_date)
        if parsed_date is None:
            return False, "Invalid date format. Use YYYY-MM-DD."
        
        cmd.extend(["/SC", "ONCE", "/SD", parsed_date.strftime("%m/%d/%Y"), "/ST", schedule_time])
        
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode != 0:
            return False, (result.stderr or result.stdout or "Task creation failed.").strip()
        return True, f"One-time cleanup scheduled at {schedule_date} {schedule_time} for {target_label}."
    
    elif schedule_mode == "interval":
        total_minutes = int(schedule.get("total_minutes", "0"))
        if total_minutes <= 0:
            return False, "Interval must be greater than 0 minutes."
            
        if total_minutes <= 1439:
            cmd.extend(["/SC", "MINUTE", "/MO", str(total_minutes)])
            unit_msg = f"{total_minutes} minute(s)"
        else:
            total_hours = max(1, total_minutes // 60)
            cmd.extend(["/SC", "HOURLY", "/MO", str(total_hours)])
            unit_msg = f"{total_hours} hour(s)"
            
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode != 0:
            return False, (result.stderr or result.stdout or "Task creation failed.").strip()
            
        return True, f"Recurring cleanup scheduled every {unit_msg} for {target_label}."
        
    return False, "Unknown schedule mode."


def remove_daily_task() -> tuple[bool, str]:
    cmd = ["schtasks", "/Delete", "/TN", TASK_NAME, "/F"]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        message = (result.stderr or result.stdout or "Task delete failed.").strip()
        lowered = message.lower()
        if "cannot find" in lowered or "cannot find the file" in lowered:
            return False, "No scheduled task found to delete."
        return False, message
    return True, "Scheduled task removed."


def run_auto_organize(path_override: str | None, template: str) -> int:
    try:
        target_path = Path(path_override) if path_override else get_desktop_path()
        if not target_path.exists() or not target_path.is_dir():
            print("Auto-run error: selected path is not a valid folder.")
            return 1

        moved, skipped, total, operations = move_files_by_extension(
            target_path=target_path,
            template=template,
            excluded_extensions=set(),
            no_extension_to_shortcuts=False,
            progress_callback=None,
        )
        log_path, state_path = _write_run_artifacts(
            target_path,
            template,
            excluded_extensions=set(),
            no_extension_to_shortcuts=False,
            moved_count=moved,
            skipped_count=skipped,
            total_count=total,
            operations=operations,
        )
        print(f"Auto-run completed. Folder={target_path} Total={total} Moved={moved} Skipped={skipped}")
        print(f"Log={log_path}")
        print(f"UndoJSON={state_path}")
        return 0
    except Exception as exc:
        print(f"Auto-run error: {exc}")
        return 1


def folder_name_for_file(file_path: Path, template: str, no_extension_to_shortcuts: bool) -> str:
    suffix = file_path.suffix.lower()
    if suffix in {".lnk", ".url"}:
        return "unknowns"
    if not suffix and no_extension_to_shortcuts:
        return "unknowns"

    extension_name = suffix.lstrip(".") if suffix else "no_extension"
    folder_name = template.replace("@", extension_name)
    return _sanitize_folder_name(folder_name)


def move_files_by_extension(
    target_path: Path,
    template: str,
    excluded_extensions: set[str],
    no_extension_to_shortcuts: bool,
    progress_callback=None,
) -> tuple[int, int, int, list[dict[str, str]]]:
    moved = 0
    skipped = 0
    operations: list[dict[str, str]] = []

    try:
        exec_path = Path(sys.executable).resolve()
    except Exception:
        exec_path = None
        
    try:
        script_path = Path(__file__).resolve()
    except Exception:
        script_path = None

    # Snapshot list prevents iterator issues while moving files.
    items = [item for item in target_path.iterdir() if item.is_file()]
    total = len(items)

    for index, item in enumerate(items, start=1):
        try:
            item_resolved = item.resolve()
        except Exception:
            item_resolved = item

        if item_resolved == exec_path or item_resolved == script_path or item.name.lower() == "desktoporganizer.exe":
            skipped += 1
            if progress_callback:
                progress_callback(index, total)
            continue

        if item.suffix.lower() in excluded_extensions:
            skipped += 1
            if progress_callback:
                progress_callback(index, total)
            continue

        target_folder = target_path / folder_name_for_file(item, template, no_extension_to_shortcuts)
        target_folder.mkdir(exist_ok=True)

        destination = target_folder / item.name
        if destination.exists():
            stem = item.stem
            suffix = item.suffix
            counter = 1
            while True:
                new_name = f"{stem}_{counter}{suffix}"
                candidate = target_folder / new_name
                if not candidate.exists():
                    destination = candidate
                    break
                counter += 1

        try:
            shutil.move(str(item), str(destination))
            moved += 1
            operations.append({"src": str(item), "dst": str(destination)})
        except Exception:
            skipped += 1

        if progress_callback:
            progress_callback(index, total)

    return moved, skipped, total, operations


class OrganizerApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Desktop Organizer")
        self.root.geometry("920x760")
        self.root.resizable(False, False)

        self.result_queue: queue.Queue = queue.Queue()
        self.worker_thread: threading.Thread | None = None

        self.default_path = self._safe_default_path()

        self.path_var = tk.StringVar(value=self.default_path)
        self.template_var = tk.StringVar(value="@_files")
        self.ignore_var = tk.StringVar(value="")
        self.no_ext_to_shortcuts_var = tk.BooleanVar(value=False)
        self.undo_json_var = tk.StringVar(value="")
        self.schedule_time_var = tk.StringVar(value="18:00")
        self.status_var = tk.StringVar(value="Ready")
        
        self.is_repeating_var = tk.BooleanVar(value=False)
        self.once_date_var = tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        self.once_time_var = tk.StringVar(value="18:00")
        self.rep_minute_var = tk.StringVar(value="0")
        self.rep_hour_var = tk.StringVar(value="0")
        self.rep_day_var = tk.StringVar(value="0")
        self.rep_week_var = tk.StringVar(value="0")
        self.rep_month_var = tk.StringVar(value="0")

        self.progress_var = tk.DoubleVar(value=0)
        self.path_var.trace_add("write", self._on_path_changed)
        self.undo_json_var.trace_add("write", self._on_path_changed)

        self.start_btn: ttk.Button | None = None
        self.undo_btn: ttk.Button | None = None
        self.progress_bar: ttk.Progressbar | None = None

        self._build_ui()
        self._refresh_undo_button_state()
        self.root.after(120, self._poll_queue)

    def _safe_default_path(self) -> str:
        try:
            return str(get_desktop_path())
        except Exception:
            return ""

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root)
        container.pack(fill="both", expand=True, padx=14, pady=14)

        title_label = ttk.Label(
            container,
            text="Desktop Organizer",
            font=("Segoe UI", 10),
        )
        title_label.grid(row=0, column=0, columnspan=3, sticky="w", pady=(10, 14), padx=12)

        path_label = ttk.Label(
            container,
            text="Start: just select a folder to organize (if empty, Desktop is used):",
            font=("Segoe UI", 10),
        )
        path_label.grid(row=1, column=0, columnspan=3, sticky="w", padx=12)

        path_entry = ttk.Entry(container, textvariable=self.path_var)
        path_entry.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(6, 10), padx=(12, 6))

        browse_btn = ttk.Button(container, text="Browse", command=self._browse_folder)
        browse_btn.grid(row=2, column=2, sticky="ew", padx=(6, 12), pady=(6, 10))

        template_label = ttk.Label(
            container,
            text="Folder name template. Use @ where extension name should be inserted:",
            font=("Segoe UI", 10),
        )
        template_label.grid(row=3, column=0, columnspan=3, sticky="w", padx=12)

        template_entry = ttk.Entry(container, textvariable=self.template_var)
        template_entry.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(6, 3), padx=12)

        template_hint = ttk.Label(
            container,
            text="Example: files_@  -> txt goes to files_txt, png goes to files_png",
            font=("Segoe UI", 10),
        )
        template_hint.grid(row=5, column=0, columnspan=3, sticky="w", pady=(0, 10), padx=12)

        ignore_label = ttk.Label(
            container,
            text="Extensions to ignore (comma separated). Please include dots, e.g. .tmp,.log,.bak:",
            font=("Segoe UI", 10),
        )
        ignore_label.grid(row=6, column=0, columnspan=3, sticky="w", padx=12)

        ignore_entry = ttk.Entry(container, textvariable=self.ignore_var)
        ignore_entry.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(6, 10), padx=12)

        no_ext_checkbox = ttk.Checkbutton(
            container,
            text="Put files without extension into unknowns (default: off)",
            variable=self.no_ext_to_shortcuts_var,
            onvalue=True,
            offvalue=False,
        )
        no_ext_checkbox.grid(row=8, column=0, columnspan=3, sticky="w", pady=(0, 10), padx=12)

        undo_json_label = ttk.Label(
            container,
            text="For Undo, find/select an undo JSON file (all file types are visible):",
            font=("Segoe UI", 10),
        )
        undo_json_label.grid(row=9, column=0, columnspan=3, sticky="w", padx=12)

        undo_json_entry = ttk.Entry(container, textvariable=self.undo_json_var)
        undo_json_entry.grid(row=10, column=0, columnspan=2, sticky="ew", pady=(6, 10), padx=(12, 6))

        browse_undo_btn = ttk.Button(container, text="Browse Undo JSON", command=self._browse_undo_json)
        browse_undo_btn.grid(row=10, column=2, sticky="ew", padx=(6, 12), pady=(6, 10))

        actions = ttk.Frame(container)
        actions.grid(row=11, column=0, columnspan=3, sticky="w", padx=12)

        self.start_btn = ttk.Button(actions, text="Start", command=self._on_start, width=15)
        self.start_btn.grid(row=0, column=0, sticky="w")

        self.undo_btn = ttk.Button(actions, text="Undo", command=self._on_undo, width=15)
        self.undo_btn.grid(row=0, column=1, sticky="w", padx=(8, 0))

        schedule_label = ttk.Label(
            container,
            text="Task Scheduler: Create automated background tasks",
            font=("Segoe UI", 10),
        )
        schedule_label.grid(row=12, column=0, columnspan=3, sticky="w", pady=(10, 0), padx=12)

        schedule_frame = ttk.Frame(container)
        schedule_frame.grid(row=13, column=0, columnspan=3, sticky="w", pady=(6, 8), padx=12)

        rep_checkbox = ttk.Checkbutton(
            schedule_frame,
            text="Repeating Schedule - Set: Minute, Hour, Day, Week, Month",
            variable=self.is_repeating_var,
            command=self._on_rep_checkbox_changed
        )
        rep_checkbox.grid(row=0, column=0, columnspan=2, sticky="nw", pady=(0, 10))

        # Box for inputs
        self.inputs_frame = ttk.Frame(schedule_frame)
        self.inputs_frame.grid(row=1, column=0, columnspan=2, sticky="w")

        # ONE-TIME
        self.once_frame = ttk.Frame(self.inputs_frame)
        self.once_frame.grid(row=0, column=0, sticky="nw", padx=(0, 10))
        ttk.Label(self.once_frame, text="One-Time").grid(row=0, column=0, columnspan=2, pady=5)
        ttk.Label(self.once_frame, text="Date (YYYY-MM-DD):").grid(row=1, column=0, padx=5, pady=(0, 5))
        self.once_date_entry = ttk.Entry(self.once_frame, textvariable=self.once_date_var, width=15)
        self.once_date_entry.grid(row=1, column=1, padx=5, pady=(0, 5))
        ttk.Label(self.once_frame, text="Time (HH:MM):").grid(row=2, column=0, padx=5, pady=(0, 5))
        self.once_time_entry = ttk.Entry(self.once_frame, textvariable=self.once_time_var, width=15)
        self.once_time_entry.grid(row=2, column=1, padx=5, pady=(0, 5))

        # REPEATING
        self.rep_frame = ttk.Frame(self.inputs_frame)
        self.rep_frame.grid(row=0, column=1, sticky="nw")
        
        ttk.Label(self.rep_frame, text="Repeating Interval").grid(row=0, column=0, columnspan=4, pady=5)
        ttk.Label(self.rep_frame, text="Min:").grid(row=1, column=0, padx=5, pady=2, sticky="e")
        ttk.Entry(self.rep_frame, textvariable=self.rep_minute_var, width=6).grid(row=1, column=1, padx=5, pady=2, sticky="w")
        ttk.Label(self.rep_frame, text="Hour:").grid(row=1, column=2, padx=5, pady=2, sticky="e")
        ttk.Entry(self.rep_frame, textvariable=self.rep_hour_var, width=6).grid(row=1, column=3, padx=5, pady=2, sticky="w")
        
        ttk.Label(self.rep_frame, text="Day:").grid(row=2, column=0, padx=5, pady=2, sticky="e")
        ttk.Entry(self.rep_frame, textvariable=self.rep_day_var, width=6).grid(row=2, column=1, padx=5, pady=2, sticky="w")
        ttk.Label(self.rep_frame, text="Week:").grid(row=2, column=2, padx=5, pady=2, sticky="e")
        ttk.Entry(self.rep_frame, textvariable=self.rep_week_var, width=6).grid(row=2, column=3, padx=5, pady=2, sticky="w")
        
        ttk.Label(self.rep_frame, text="Month:").grid(row=3, column=0, padx=5, pady=2, sticky="e")
        ttk.Entry(self.rep_frame, textvariable=self.rep_month_var, width=6).grid(row=3, column=1, padx=5, pady=2, sticky="w")

        btn_container = ttk.Frame(schedule_frame)
        btn_container.grid(row=2, column=0, sticky="w", pady=(10, 0))

        schedule_create_btn = ttk.Button(btn_container, text="Create Scheduled Task", command=self._on_create_schedule, width=25)
        schedule_create_btn.grid(row=0, column=0, sticky="w")

        schedule_remove_btn = ttk.Button(btn_container, text="Remove Task", command=self._on_remove_schedule, width=15)
        schedule_remove_btn.grid(row=0, column=1, sticky="w", padx=(8, 0))

        self.progress_bar = ttk.Progressbar(container)
        self.progress_bar.grid(row=14, column=0, columnspan=3, sticky="ew", pady=(12, 6), padx=12)
        self.progress_bar['value'] = 0

        status_label = ttk.Label(container, textvariable=self.status_var, font=("Segoe UI", 10))
        status_label.grid(row=15, column=0, columnspan=3, sticky="w", padx=12, pady=(0, 10))

        container.columnconfigure(0, weight=1)
        container.columnconfigure(1, weight=1)
        container.columnconfigure(2, weight=0)

        self._on_rep_checkbox_changed()

    def _on_rep_checkbox_changed(self) -> None:
        if self.is_repeating_var.get():
            # Disable Once, Enable Repeating
            for child in self.once_frame.winfo_children():
                try: child.configure(state="disabled")
                except: pass
            for child in self.rep_frame.winfo_children():
                try: child.configure(state="normal")
                except: pass
        else:
            # Enable Once, Disable Repeating
            for child in self.once_frame.winfo_children():
                try: child.configure(state="normal")
                except: pass
            for child in self.rep_frame.winfo_children():
                try: child.configure(state="disabled")
                except: pass

    def _browse_folder(self) -> None:
        selected = filedialog.askdirectory(title="Select folder to organize")
        if selected:
            self.path_var.set(selected)

    def _browse_undo_json(self) -> None:
        selected = filedialog.askopenfilename(
            title="Select undo JSON file",
            filetypes=[("Undo JSON", "*.json"), ("All files", "*.*")],
        )
        if selected:
            self.undo_json_var.set(selected)

    def _resolve_current_path(self) -> Path | None:
        path_text = self.path_var.get().strip()
        if not path_text:
            try:
                path_text = str(get_desktop_path())
                self.path_var.set(path_text)
            except Exception:
                return None

        selected_path = Path(path_text)
        if not selected_path.exists() or not selected_path.is_dir():
            return None
        return selected_path

    def _on_path_changed(self, *_args) -> None:
        self._refresh_undo_button_state()

    def _refresh_undo_button_state(self) -> None:
        json_path_text = self.undo_json_var.get().strip()
        if json_path_text and Path(json_path_text).exists():
            if self.undo_btn is not None:
                self.undo_btn.configure(state="normal")
            return

        path = self._resolve_current_path()
        if path and can_undo_for_path(path):
            if self.undo_btn is not None:
                self.undo_btn.configure(state="normal")
        else:
            if self.undo_btn is not None:
                self.undo_btn.configure(state="disabled")

    def _set_running_state(self, is_running: bool) -> None:
        if is_running:
            if self.start_btn is not None:
                self.start_btn.configure(state="disabled")
            if self.undo_btn is not None:
                self.undo_btn.configure(state="disabled")
        else:
            if self.start_btn is not None:
                self.start_btn.configure(state="normal")
            self._refresh_undo_button_state()

    def _on_start(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            return

        template = self.template_var.get().strip()
        if not template:
            messagebox.showerror("Validation", "Template cannot be empty.")
            return

        if "@" not in template:
            messagebox.showerror("Validation", "Template must include @ placeholder.")
            return

        selected_path = self._resolve_current_path()
        if selected_path is None:
            messagebox.showerror("Validation", "Selected path is not a valid folder.")
            return

        excluded_extensions, missing_dot_entries = _parse_excluded_extensions(self.ignore_var.get())
        if missing_dot_entries:
            messagebox.showinfo(
                "Note",
                "Some ignore extensions were entered without dot. "
                "They were auto-normalized: "
                + ", ".join(f".{entry}" for entry in missing_dot_entries),
            )

        self.progress_var.set(0)
        if self.progress_bar is not None:
            self.progress_bar['value'] = 0
        self.status_var.set("Processing...")
        self._set_running_state(True)

        self.worker_thread = threading.Thread(
            target=self._run_organizer,
            args=(
                selected_path,
                template,
                excluded_extensions,
                self.no_ext_to_shortcuts_var.get(),
            ),
            daemon=True,
        )
        self.worker_thread.start()

    def _on_undo(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            return

        selected_path = self._resolve_current_path()
        json_path_text = self.undo_json_var.get().strip()
        undo_json_path: Path | None = None

        if json_path_text:
            candidate = Path(json_path_text)
            if not candidate.exists() or candidate.suffix.lower() != ".json":
                messagebox.showerror("Validation", "Please select a valid undo JSON file.")
                return
            undo_json_path = candidate
        else:
            if selected_path is None:
                messagebox.showerror("Undo", "For undo, please find/select an undo JSON file.")
                return
            undo_json_path = _latest_undo_file_for_folder(selected_path)
            if undo_json_path is None:
                messagebox.showinfo("Undo", "For undo, find/select an undo JSON file.")
                return
            self.undo_json_var.set(str(undo_json_path))

        confirmed = messagebox.askyesno(
            "Undo",
            "Undo will try to move files back to their previous locations from the last run. Continue?",
        )
        if not confirmed:
            return

        self.progress_var.set(0)
        if self.progress_bar is not None:
            self.progress_bar['value'] = 0
        self.status_var.set("Undo in progress...")
        self._set_running_state(True)

        self.worker_thread = threading.Thread(
            target=self._run_undo,
            args=(undo_json_path,),
            daemon=True,
        )
        self.worker_thread.start()

    def _on_create_schedule(self) -> None:
        schedule: dict[str, str] = {}
        
        if not self.is_repeating_var.get():
            # One-time task
            date_text = self.once_date_var.get().strip()
            time_text = self.once_time_var.get().strip()
            
            if _parse_iso_date(date_text) is None:
                messagebox.showerror("Validation", "Invalid date format. Use YYYY-MM-DD.")
                return
            if not _is_valid_hhmm(time_text):
                messagebox.showerror("Validation", "Invalid time format. Use HH:MM.")
                return
                
            schedule = {
                "mode": "once",
                "date": date_text,
                "time": time_text
            }
        else:
            # Interval task
            try:
                m = float(self.rep_minute_var.get().strip() or "0")
                h = float(self.rep_hour_var.get().strip() or "0")
                d = float(self.rep_day_var.get().strip() or "0")
                w = float(self.rep_week_var.get().strip() or "0")
                mo = float(self.rep_month_var.get().strip() or "0")
            except ValueError:
                messagebox.showerror("Validation", "Please enter numeric values only.")
                return
                
            total_minutes = int(m + (h * 60) + (d * 24 * 60) + (w * 7 * 24 * 60) + (mo * 30.416 * 24 * 60))
            if total_minutes <= 0:
                messagebox.showerror("Validation", "Repeating interval must be greater than 0.")
                return
                
            schedule = {
                "mode": "interval",
                "total_minutes": str(total_minutes)
            }

        path = self._resolve_current_path()
        if path is None:
            path = None

        ok, message = create_or_update_task(schedule, path)
        self.status_var.set(message)
        if ok:
            messagebox.showinfo("Task Scheduler", message)
        else:
            messagebox.showerror("Task Scheduler", message)

    def _on_remove_schedule(self) -> None:
        ok, message = remove_daily_task()
        self.status_var.set(message)
        if ok:
            messagebox.showinfo("Task Scheduler", message)
        else:
            messagebox.showerror("Task Scheduler", message)

    def _run_organizer(
        self,
        selected_path: Path,
        template: str,
        excluded_extensions: set[str],
        no_extension_to_shortcuts: bool,
    ) -> None:
        def on_progress(current: int, total: int) -> None:
            percent = 100 if total == 0 else (current / total) * 100
            self.result_queue.put(("progress", current, total, percent))

        try:
            moved, skipped, total, operations = move_files_by_extension(
                selected_path,
                template,
                excluded_extensions,
                no_extension_to_shortcuts,
                progress_callback=on_progress,
            )
            log_path, _ = _write_run_artifacts(
                selected_path,
                template,
                excluded_extensions,
                no_extension_to_shortcuts,
                moved,
                skipped,
                total,
                operations,
            )
            undo_json_path = _latest_undo_file_for_folder(selected_path)
            self.result_queue.put(("done", moved, skipped, total, str(selected_path)))
            self.result_queue.put(("log", str(log_path)))
            if undo_json_path is not None:
                self.result_queue.put(("undo_json", str(undo_json_path)))
        except Exception as exc:
            self.result_queue.put(("error", str(exc)))

    def _run_undo(self, undo_json_path: Path) -> None:
        def on_progress(current: int, total: int) -> None:
            percent = 100 if total == 0 else (current / total) * 100
            self.result_queue.put(("undo_progress", current, total, percent))

        try:
            undone, skipped, total, target_folder = undo_last_run(undo_json_path, progress_callback=on_progress)
            self.result_queue.put(("undo_done", undone, skipped, total, target_folder, str(undo_json_path)))
        except Exception as exc:
            self.result_queue.put(("undo_error", str(exc)))

    def _poll_queue(self) -> None:
        while True:
            try:
                item = self.result_queue.get_nowait()
            except queue.Empty:
                break

            kind = item[0]
            if kind == "progress":
                _, current, total, percent = item
                self.progress_var.set(percent)
                if self.progress_bar is not None:
                    self.progress_bar['value'] = percent
                self.status_var.set(f"Processing... {current}/{total}")
            elif kind == "undo_progress":
                _, current, total, percent = item
                self.progress_var.set(percent)
                if self.progress_bar is not None:
                    self.progress_bar['value'] = percent
                self.status_var.set(f"Undo in progress... {current}/{total}")
            elif kind == "done":
                _, moved, skipped, total, path_text = item
                self.progress_var.set(100)
                if self.progress_bar is not None:
                    self.progress_bar['value'] = 100
                self.status_var.set(
                    f"Completed. Folder: {path_text} | Total: {total}, Moved: {moved}, Skipped: {skipped}"
                )
                self._set_running_state(False)
                messagebox.showinfo(
                    "Completed",
                    f"Completed.\n\nFolder: {path_text}\nTotal files: {total}\nMoved: {moved}\nSkipped: {skipped}",
                )
            elif kind == "undo_done":
                _, undone, skipped, total, path_text, undo_json = item
                self.progress_var.set(100)
                if self.progress_bar is not None:
                    self.progress_bar['value'] = 100
                self.status_var.set(
                    f"Undo completed. Folder: {path_text} | Total: {total}, Restored: {undone}, Skipped: {skipped}"
                )
                self._set_running_state(False)
                self.undo_json_var.set("")
                messagebox.showinfo(
                    "Undo Completed",
                    (
                        f"Undo completed.\n\nFolder: {path_text}\n"
                        f"Tracked files: {total}\nRestored: {undone}\nSkipped: {skipped}\nUndo JSON: {undo_json}"
                    ),
                )
            elif kind == "error":
                _, error_text = item
                self.status_var.set("Error occurred.")
                self._set_running_state(False)
                messagebox.showerror("Error", error_text)
            elif kind == "undo_error":
                _, error_text = item
                self.status_var.set("Undo error occurred.")
                self._set_running_state(False)
                messagebox.showerror("Undo Error", error_text)
            elif kind == "log":
                _, log_path = item
                self.status_var.set(f"Log saved: {log_path}")
            elif kind == "undo_json":
                _, undo_json_path = item
                self.undo_json_var.set(undo_json_path)
                self.status_var.set(f"Undo JSON saved: {undo_json_path}")

        self.root.after(120, self._poll_queue)


def main() -> int:
    parser = argparse.ArgumentParser(add_help=True)
    parser.add_argument("--auto-run", action="store_true", help="Run organizer without GUI.")
    parser.add_argument("--path", default=None, help="Optional target folder for auto-run.")
    parser.add_argument("--template", default="@_files", help="Folder naming template for auto-run.")
    args = parser.parse_args()

    if args.auto_run:
        return run_auto_organize(path_override=args.path, template=args.template)

    app_root = tk.Tk()
    OrganizerApp(app_root)
    app_root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
