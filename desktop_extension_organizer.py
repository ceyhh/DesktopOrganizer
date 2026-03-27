import ctypes
import os
import queue
import shutil
import sys
import threading
import tkinter as tk
import unicodedata
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

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
) -> tuple[int, int, int]:
    moved = 0
    skipped = 0

    # Snapshot list prevents iterator issues while moving files.
    items = [item for item in target_path.iterdir() if item.is_file()]
    total = len(items)

    for index, item in enumerate(items, start=1):
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
        except Exception:
            skipped += 1

        if progress_callback:
            progress_callback(index, total)

    return moved, skipped, total


class OrganizerApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Desktop Organizer")
        self.root.geometry("700x390")
        self.root.resizable(False, False)

        self.result_queue: queue.Queue = queue.Queue()
        self.worker_thread: threading.Thread | None = None

        self.default_path = self._safe_default_path()

        self.path_var = tk.StringVar(value=self.default_path)
        self.template_var = tk.StringVar(value="@_files")
        self.ignore_var = tk.StringVar(value="")
        self.no_ext_to_shortcuts_var = tk.BooleanVar(value=False)
        self.status_var = tk.StringVar(value="Ready")
        self.progress_var = tk.DoubleVar(value=0)

        self._build_ui()
        self.root.after(120, self._poll_queue)

    def _safe_default_path(self) -> str:
        try:
            return str(get_desktop_path())
        except Exception:
            return ""

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=14)
        container.pack(fill="both", expand=True)

        title_label = ttk.Label(container, text="File Organizer", font=("Segoe UI", 13, "bold"))
        title_label.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 10))

        path_label = ttk.Label(container, text="Folder path to organize (if empty, Desktop is used):")
        path_label.grid(row=1, column=0, columnspan=3, sticky="w")

        path_entry = ttk.Entry(container, textvariable=self.path_var, width=75)
        path_entry.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(4, 8))

        browse_btn = ttk.Button(container, text="Browse", command=self._browse_folder)
        browse_btn.grid(row=2, column=2, sticky="ew", padx=(8, 0), pady=(4, 8))

        template_label = ttk.Label(
            container,
            text="Folder name template. Use @ where extension name should be inserted:",
        )
        template_label.grid(row=3, column=0, columnspan=3, sticky="w")

        template_entry = ttk.Entry(container, textvariable=self.template_var, width=75)
        template_entry.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(4, 2))

        template_hint = ttk.Label(
            container,
            text="Example: files_@  -> txt goes to files_txt, png goes to files_png",
        )
        template_hint.grid(row=5, column=0, columnspan=3, sticky="w", pady=(0, 8))

        ignore_label = ttk.Label(
            container,
            text="Extensions to ignore (comma separated). Please include dots, e.g. .tmp,.log,.bak:",
        )
        ignore_label.grid(row=6, column=0, columnspan=3, sticky="w")

        ignore_entry = ttk.Entry(container, textvariable=self.ignore_var, width=75)
        ignore_entry.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(4, 8))

        no_ext_checkbox = ttk.Checkbutton(
            container,
            text="Put files without extension into unknowns (default: off)",
            variable=self.no_ext_to_shortcuts_var,
        )
        no_ext_checkbox.grid(row=8, column=0, columnspan=3, sticky="w", pady=(0, 8))

        self.start_btn = ttk.Button(container, text="Start", command=self._on_start)
        self.start_btn.grid(row=9, column=0, sticky="w")

        progress_bar = ttk.Progressbar(
            container,
            variable=self.progress_var,
            maximum=100,
            mode="determinate",
        )
        progress_bar.grid(row=10, column=0, columnspan=3, sticky="ew", pady=(10, 6))

        status_label = ttk.Label(container, textvariable=self.status_var)
        status_label.grid(row=11, column=0, columnspan=3, sticky="w")

        container.columnconfigure(0, weight=1)
        container.columnconfigure(1, weight=1)
        container.columnconfigure(2, weight=0)

    def _browse_folder(self) -> None:
        selected = filedialog.askdirectory(title="Select folder to organize")
        if selected:
            self.path_var.set(selected)

    def _set_running_state(self, is_running: bool) -> None:
        if is_running:
            self.start_btn.state(["disabled"])
        else:
            self.start_btn.state(["!disabled"])

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

        path_text = self.path_var.get().strip()
        if not path_text:
            try:
                path_text = str(get_desktop_path())
                self.path_var.set(path_text)
            except Exception as exc:
                messagebox.showerror("Desktop Error", f"Desktop path could not be found: {exc}")
                return

        selected_path = Path(path_text)
        if not selected_path.exists() or not selected_path.is_dir():
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
            moved, skipped, total = move_files_by_extension(
                selected_path,
                template,
                excluded_extensions,
                no_extension_to_shortcuts,
                progress_callback=on_progress,
            )
            self.result_queue.put(("done", moved, skipped, total, str(selected_path)))
        except Exception as exc:
            self.result_queue.put(("error", str(exc)))

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
                self.status_var.set(f"Processing... {current}/{total}")
            elif kind == "done":
                _, moved, skipped, total, path_text = item
                self.progress_var.set(100)
                self.status_var.set(
                    f"Tamamlandi. Folder: {path_text} | Total: {total}, Moved: {moved}, Skipped: {skipped}"
                )
                self._set_running_state(False)
                messagebox.showinfo(
                    "Tamamlandi",
                    f"Tamamlandi.\n\nFolder: {path_text}\nTotal files: {total}\nMoved: {moved}\nSkipped: {skipped}",
                )
            elif kind == "error":
                _, error_text = item
                self.status_var.set("Error occurred.")
                self._set_running_state(False)
                messagebox.showerror("Error", error_text)

        self.root.after(120, self._poll_queue)


def main() -> int:
    app_root = tk.Tk()
    OrganizerApp(app_root)
    app_root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
