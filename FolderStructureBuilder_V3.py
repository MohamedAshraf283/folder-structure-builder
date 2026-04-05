#!/usr/bin/env python3
# -*- coding: utf-8 -*-

# ==============================================================================
# Folder Structure Builder  v3  —  Professional Dark Edition
# ==============================================================================
# New in v3:
#   + Dark professional theme with strong accent colours
#   + Coloured log (SUCCESS / ERROR / WARNING / INFO / SECTION tags)
#   + Status bar (job count, last action)
#   + Export / Import job list as JSON
#   + Duplicate selected job
#   + Progress bar during "Run All"
#   + "Copy Path" button in Dry Run window
#   + Tooltips on all buttons
#   + Resizable Job dialog
#   + Recursive symlink detection anywhere in source tree
#   + Dry-run works even when target dir doesn't exist yet
#   + Preview Plan now shows a real TreeView (same as Dry Run window)
#
# Bug fixes vs v2:
#   * _sanitize: double-backslash replaced correctly (/ -> \ not \\)
# ==============================================================================

from __future__ import annotations

import json
import os
import shutil
from dataclasses import dataclass
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox


# ==============================================================================
# Colour palette
# ==============================================================================
_C: dict[str, str] = {
    "bg":        "#13141D",   # window background
    "surface":   "#1A1B28",   # card / labelframe
    "surface2":  "#21223A",   # row alternation / hover
    "surface3":  "#2A2C45",   # inputs, trough, buttons
    "accent":    "#7B68EE",   # primary — medium-slate-blue
    "accent_h":  "#9D8FFF",   # accent hover
    "accent2":   "#00CFFF",   # secondary — cyan
    "success":   "#3DDC84",   # green
    "error":     "#FF5370",   # red
    "warning":   "#FFB74D",   # amber
    "text":      "#DDE1F0",   # primary text
    "text_dim":  "#6E7490",   # muted text
    "border":    "#292A40",   # borders / separators
    "hdr":       "#0B0C14",   # header bar
    "sel":       "#7B68EE",   # treeview selection (same as accent)
}


# ==============================================================================
# Dark theme
# ==============================================================================
def _apply_theme(root: tk.Tk) -> ttk.Style:
    s = ttk.Style(root)
    s.theme_use("clam")

    BG   = _C["bg"]
    SURF = _C["surface"]
    S2   = _C["surface2"]
    S3   = _C["surface3"]
    ACC  = _C["accent"]
    TXT  = _C["text"]
    DIM  = _C["text_dim"]
    BRD  = _C["border"]

    root.configure(bg=BG)

    # Base defaults
    s.configure(".",
        background=BG, foreground=TXT,
        fieldbackground=S3,
        troughcolor=S2,
        bordercolor=BRD, darkcolor=BRD, lightcolor=BRD,
        selectbackground=ACC, selectforeground="#FFFFFF",
        relief="flat", font=("Segoe UI", 10),
    )

    # Frames
    s.configure("TFrame",        background=BG)
    s.configure("Card.TFrame",   background=SURF)
    s.configure("Header.TFrame", background=_C["hdr"])

    # Labels
    s.configure("TLabel",      background=BG,   foreground=TXT)
    s.configure("Card.TLabel", background=SURF, foreground=TXT)
    s.configure("Dim.TLabel",  background=BG,   foreground=DIM, font=("Segoe UI", 9))

    # LabelFrame
    s.configure("TLabelframe",
        background=SURF, bordercolor=BRD, relief="groove")
    s.configure("TLabelframe.Label",
        background=SURF, foreground=ACC, font=("Segoe UI", 10, "bold"))

    # ── Buttons ───────────────────────────────────────────────────────────────
    # Default
    s.configure("TButton",
        background=S3, foreground=TXT,
        bordercolor=BRD, focuscolor=ACC,
        relief="flat", padding=(10, 5),
        font=("Segoe UI", 10),
    )
    s.map("TButton",
        background=[("active", S2), ("pressed", ACC)],
        foreground=[("active", TXT), ("pressed", "#fff")],
        bordercolor=[("active", ACC)],
    )

    # Primary / Accent
    s.configure("Accent.TButton",
        background=ACC, foreground="#ffffff",
        bordercolor=ACC, relief="flat",
        padding=(14, 7), font=("Segoe UI", 10, "bold"),
    )
    s.map("Accent.TButton",
        background=[("active", _C["accent_h"]), ("pressed", _C["accent_h"])],
        bordercolor=[("active", _C["accent_h"])],
    )

    # Danger / destructive
    s.configure("Danger.TButton",
        background="#220F15", foreground=_C["error"],
        bordercolor="#3D1520", relief="flat",
        padding=(10, 5), font=("Segoe UI", 10),
    )
    s.map("Danger.TButton",
        background=[("active", "#321520")],
        bordercolor=[("active", _C["error"])],
    )

    # Success
    s.configure("Success.TButton",
        background="#0E2218", foreground=_C["success"],
        bordercolor="#1E3D28", relief="flat",
        padding=(10, 5), font=("Segoe UI", 10),
    )
    s.map("Success.TButton",
        background=[("active", "#182D20")],
        bordercolor=[("active", _C["success"])],
    )

    # Entry
    s.configure("TEntry",
        fieldbackground=S3, foreground=TXT,
        insertcolor=_C["accent2"],
        bordercolor=BRD, lightcolor=BRD, darkcolor=BRD,
        relief="flat", padding=(7, 5),
    )
    s.map("TEntry",
        bordercolor=[("focus", ACC)],
        lightcolor=[("focus", ACC)],
        darkcolor=[("focus", ACC)],
    )

    # Checkbutton
    s.configure("TCheckbutton",
        background=BG, foreground=TXT, indicatorcolor=S3)
    s.map("TCheckbutton",
        background=[("active", BG)],
        indicatorcolor=[("selected", ACC), ("active", S2)],
        foreground=[("active", TXT)],
    )

    # Combobox
    s.configure("TCombobox",
        fieldbackground=S3, background=S2,
        foreground=TXT, arrowcolor=DIM,
        bordercolor=BRD, relief="flat",
    )
    s.map("TCombobox",
        fieldbackground=[("readonly", S3)],
        background=[("active", S3)],
        bordercolor=[("focus", ACC)],
    )

    # Treeview
    s.configure("Treeview",
        background=SURF, foreground=TXT,
        fieldbackground=SURF, rowheight=30,
        bordercolor=BRD, relief="flat",
        font=("Segoe UI", 10),
    )
    s.configure("Treeview.Heading",
        background=S2, foreground=DIM,
        relief="flat", font=("Segoe UI", 10, "bold"),
    )
    s.map("Treeview",
        background=[("selected", ACC)],
        foreground=[("selected", "#ffffff")],
    )
    s.map("Treeview.Heading",
        background=[("active", S3)],
        foreground=[("active", TXT)],
    )

    # Scrollbar
    s.configure("TScrollbar",
        background=S2, troughcolor=SURF,
        bordercolor=BRD, arrowcolor=DIM,
        relief="flat", arrowsize=12,
    )
    s.map("TScrollbar", background=[("active", S3)])

    # Progressbar
    s.configure("TProgressbar",
        background=ACC, troughcolor=S2,
        bordercolor=BRD, thickness=6,
    )

    # Separator
    s.configure("TSeparator", background=BRD)

    return s


# ==============================================================================
# Tooltip
# ==============================================================================
class _Tooltip:
    """Simple hover tooltip for any tkinter widget."""

    def __init__(self, widget: tk.Widget, text: str) -> None:
        self._w    = widget
        self._text = text
        self._win: tk.Toplevel | None = None
        widget.bind("<Enter>",       self._show, add="+")
        widget.bind("<Leave>",       self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")

    def _show(self, _e=None) -> None:
        if self._win:
            return
        x = self._w.winfo_rootx() + 14
        y = self._w.winfo_rooty() + self._w.winfo_height() + 6
        self._win = tw = tk.Toplevel(self._w)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tk.Label(
            tw, text=self._text,
            bg=_C["surface2"], fg=_C["text"],
            relief="flat", bd=1,
            font=("Segoe UI", 9),
            padx=10, pady=5,
        ).pack()

    def _hide(self, _e=None) -> None:
        if self._win:
            self._win.destroy()
            self._win = None


# ==============================================================================
# Clipboard helper (with path sanitisation)
# ==============================================================================
def _bind_clipboard(widget: tk.Widget) -> None:
    """
    Explicit clipboard bindings that survive grab_set() on Windows.
    Supports multiple keyboard layouts (Arabic, etc.) via VK-code fallback.
    Pasted text is sanitised for use as a Windows path.
    Right-click menu included.
    """

    def _sanitize(text: str) -> str:
        # Strip newlines, outer quotes, and extra spaces.
        text = text.replace("\r", "").replace("\n", "")
        text = text.strip().strip('"').strip("'").strip()
        # BUG FIX v3: replace forward-slashes with ONE backslash (was \\).
        text = text.replace("/", "\\")
        return text

    def _drop_sel() -> None:
        try:
            widget.delete(  # type: ignore[attr-defined]
                widget.index(tk.SEL_FIRST),  # type: ignore[attr-defined]
                widget.index(tk.SEL_LAST),   # type: ignore[attr-defined]
            )
        except (tk.TclError, AttributeError):
            pass

    def paste(_e=None) -> str:
        try:
            text = widget.clipboard_get()
        except tk.TclError:
            return "break"
        text = _sanitize(text)
        _drop_sel()
        try:
            widget.insert(tk.INSERT, text)  # type: ignore[attr-defined]
        except Exception:
            pass
        return "break"

    def copy_sel(_e=None) -> str:
        try:
            t = widget.selection_get()
            widget.clipboard_clear()
            widget.clipboard_append(t)
        except (tk.TclError, AttributeError):
            pass
        return "break"

    def cut(_e=None) -> str:
        try:
            t = widget.selection_get()
            widget.clipboard_clear()
            widget.clipboard_append(t)
            widget.delete(tk.SEL_FIRST, tk.SEL_LAST)  # type: ignore[attr-defined]
        except (tk.TclError, AttributeError):
            pass
        return "break"

    def sel_all(_e=None) -> str:
        try:
            widget.select_range(0, tk.END)  # type: ignore[attr-defined]
            widget.icursor(tk.END)          # type: ignore[attr-defined]
        except (AttributeError, tk.TclError):
            pass
        return "break"

    # Standard sequences
    for seq in ("<Control-v>", "<Control-V>", "<Shift-Insert>"):
        widget.bind(seq, paste)
    for seq in ("<Control-c>", "<Control-C>", "<Control-Insert>"):
        widget.bind(seq, copy_sel)
    for seq in ("<Control-x>", "<Control-X>", "<Shift-Delete>"):
        widget.bind(seq, cut)
    for seq in ("<Control-a>", "<Control-A>"):
        widget.bind(seq, sel_all)

    # VK-code fallback (keyboard-layout-independent on Windows)
    _VK = {65: sel_all, 67: copy_sel, 86: paste, 88: cut}

    def _vk(event: tk.Event) -> str | None:
        fn = _VK.get(getattr(event, "keycode", None))
        return fn(event) if fn else None

    widget.bind("<Control-KeyPress>", _vk)

    # Right-click context menu
    menu = tk.Menu(widget, tearoff=0,
                   bg=_C["surface2"], fg=_C["text"],
                   activebackground=_C["accent"], activeforeground="#fff",
                   relief="flat", bd=1)
    menu.add_command(label="  Paste",      command=paste)
    menu.add_command(label="  Copy",       command=copy_sel)
    menu.add_command(label="  Cut",        command=cut)
    menu.add_separator()
    menu.add_command(label="  Select All", command=sel_all)

    def _show_menu(ev: tk.Event) -> None:
        try:
            menu.tk_popup(ev.x_root, ev.y_root)
        finally:
            menu.grab_release()

    widget.bind("<Button-3>", _show_menu)
    widget.bind("<Button-2>", _show_menu)


# ==============================================================================
# Configuration
# ==============================================================================
CONFIG = {
    "app_title":   "Folder Structure Builder  —  by Mohamed Ashraf",
    "window_size": "1080x870",
    "min_size":    (1000, 800),
    "padding":     12,
    "version":     "v3.0",
    "defaults": {
        "allow_absolute_structure_path": False,
        "forbid_traversal":              True,
        "exist_ok":                      True,
        "dry_run":                       False,
        "disallow_symlinks":             True,
    },
}

_WINDOWS_RESERVED: set[str] = {
    "CON", "PRN", "AUX", "NUL",
    "COM1", "COM2", "COM3", "COM4", "COM5", "COM6", "COM7", "COM8", "COM9",
    "LPT1", "LPT2", "LPT3", "LPT4", "LPT5", "LPT6", "LPT7", "LPT8", "LPT9",
}

SETTINGS_HELP_TEXT = {
    "en": (
        "Settings Explanation\n\n"
        "1) Dry Run\n"
        "   Simulates all operations — no folders created, nothing copied.\n"
        "   Opens a plan window showing every path that would be affected.\n\n"
        "2) Allow absolute structure path\n"
        "   Allows paths like  C:\\xampp\\htdocs\\app\n"
        "   The drive root (C:\\) is stripped; everything goes under Output Root.\n"
        "   Example:  C:\\xampp\\htdocs  =>  <OutputRoot>\\xampp\\htdocs\n\n"
        "3) Forbid '..' traversal\n"
        "   Rejects any structure path containing '..'.\n\n"
        "4) Disallow symlinks\n"
        "   Rejects symlinks in the copy source (checked recursively).\n\n"
        "Copy rules:\n"
        "  • NO overwrite — any existing destination raises an error.\n"
        "  • Folder source — CONTENTS only (not the container folder).\n"
        "  • Mid-copy failure — already-copied items are rolled back.\n\n"
        "Export / Import:\n"
        "  Save and reload your job list as a JSON file.\n\n"
        "General:\n"
        "  Output Root must be an existing directory.\n"
        "  Windows reserved names (CON, NUL, COM1-9…) are rejected.\n"
    ),
    "ar": (
        "شرح الإعدادات\n\n"
        "1) Dry Run\n"
        "   يحاكي جميع العمليات بدون أي تغييرات فعلية.\n"
        "   يفتح نافذة تعرض كل المسارات التي كانت ستتأثر.\n\n"
        "2) Allow absolute structure path\n"
        "   يسمح بمسارات مثل  C:\\xampp\\htdocs\\app\n"
        "   يتجاهل الـ Drive Root (مثل C:\\) والنتيجة دائماً تحت Output Root.\n"
        "   مثال:  C:\\xampp\\htdocs  =>  <OutputRoot>\\xampp\\htdocs\n\n"
        "3) Forbid '..' traversal\n"
        "   يرفض أي مسار يحتوي على '..'.\n\n"
        "4) Disallow symlinks\n"
        "   يرفض الروابط الرمزية في مصدر النسخ (فحص متكرر عميق).\n\n"
        "قواعد النسخ:\n"
        "  • لا Overwrite — أي وجهة موجودة = خطأ.\n"
        "  • مصدر مجلد — المحتويات فقط بدون المجلد الحاوي.\n"
        "  • فشل في المنتصف — يتم Rollback لكل ما نُسخ.\n\n"
        "Export / Import:\n"
        "  احفظ وأعِد تحميل قائمة المهام كملف JSON.\n\n"
        "ملاحظات:\n"
        "  Output Root يجب أن يكون مجلداً موجوداً.\n"
        "  الأسماء المحجوزة في Windows (CON, NUL, COM1-9…) مرفوضة.\n"
    ),
}


# ==============================================================================
# Domain / Core Logic
# ==============================================================================
@dataclass(frozen=True)
class Job:
    structure_path: str
    copy_source: Path | None = None


class StructurePathError(ValueError):
    pass


class CopySourceError(ValueError):
    pass


def safe_resolve_user_path(
    raw: str,
    *,
    must_exist: bool = False,
    resolve: bool = True,
) -> Path:
    raw = (raw or "").strip().strip('"').strip("'").strip()
    if not raw:
        raise ValueError("Path is empty.")
    path = Path(raw).expanduser()
    if must_exist and not path.exists():
        raise ValueError(f"Path does not exist: {path}")
    return path.resolve() if resolve else Path(os.path.abspath(path))


def normalize_structure_path(
    raw: str,
    *,
    allow_absolute: bool,
    forbid_traversal: bool,
) -> list[str]:
    raw = (raw or "").strip()
    if not raw:
        raise StructurePathError("Structure Path is empty.")

    if not allow_absolute:
        if (
            raw.startswith("\\") or raw.startswith("/")
            or (len(raw) >= 2 and raw[1] == ":")
        ):
            raise StructurePathError(
                "Structure Path must be relative (no drive letter like C:\\...).\n"
                "Enable 'Allow absolute structure path' in Settings if needed."
            )

    win = raw.replace("/", "\\").strip()
    if allow_absolute and len(win) >= 2 and win[1] == ":":
        win = win[2:].lstrip("\\")

    parts = [c.strip() for c in win.split("\\") if c.strip()]
    if not parts:
        raise StructurePathError("No folder segments found in Structure Path.")

    if forbid_traversal and any(p == ".." for p in parts):
        raise StructurePathError(
            "Structure Path contains '..' (path traversal) — not allowed."
        )

    parts = [p for p in parts if p != "."]

    illegal = set('<>:"/\\|?*')
    for p in parts:
        if p.upper() in _WINDOWS_RESERVED:
            raise StructurePathError(
                f"'{p}' is a Windows reserved name and cannot be used as a folder."
            )
        if any(ch in illegal for ch in p):
            raise StructurePathError(
                f"Folder name '{p}' contains illegal characters."
            )
        if p.endswith(" ") or p.endswith("."):
            raise StructurePathError(
                f"Folder name '{p}' cannot end with a space or dot on Windows."
            )

    return parts


def compute_target_path(
    output_root: Path,
    structure_path: str,
    *,
    allow_absolute: bool,
    forbid_traversal: bool,
) -> Path:
    output_root = Path(output_root).expanduser().resolve()
    if not output_root.exists() or not output_root.is_dir():
        raise StructurePathError("Output Root must be an existing directory.")

    parts = normalize_structure_path(
        structure_path,
        allow_absolute=allow_absolute,
        forbid_traversal=forbid_traversal,
    )
    target = output_root
    for seg in parts:
        target = target / seg
    return target


def ensure_dir(path: Path, *, exist_ok: bool, dry_run: bool) -> None:
    if not dry_run:
        path.mkdir(parents=True, exist_ok=exist_ok)


def _is_symlink(path: Path) -> bool:
    try:
        return path.is_symlink()
    except Exception:
        return False


def _find_symlink_in_tree(root: Path) -> Path | None:
    """Return the first symlink found anywhere under root, or None."""
    for dirpath, dirnames, filenames in os.walk(root, followlinks=False):
        base = Path(dirpath)
        for name in list(dirnames):
            p = base / name
            if _is_symlink(p):
                return p
        for name in filenames:
            p = base / name
            if _is_symlink(p):
                return p
    return None


def copy_into_target(
    src: Path,
    target_dir: Path,
    *,
    disallow_symlinks: bool,
    dry_run: bool,
) -> tuple[list[str], list[dict]]:
    """Copy file / folder-contents into target_dir with strict no-overwrite."""
    actions: list[str] = []
    planned: list[dict] = []

    src_raw   = Path(src).expanduser()
    target_dir = Path(target_dir).expanduser().resolve()

    if not src_raw.exists():
        raise CopySourceError(f"Copy source does not exist: {src_raw}")
    if disallow_symlinks and _is_symlink(src_raw):
        raise CopySourceError(f"Symlink source is not allowed: {src_raw}")

    src = src_raw.resolve()

    if target_dir.exists():
        if not target_dir.is_dir():
            raise CopySourceError(
                f"Target exists but is not a directory: {target_dir}"
            )
    elif not dry_run:
        raise CopySourceError("Internal error: target directory does not exist.")

    # ── Single file ───────────────────────────────────────────────────────────
    if src.is_file():
        dest = target_dir / src.name
        if dest.exists():
            raise CopySourceError(f"Destination already exists: {dest}")
        planned.append({"kind": "FILE", "path": str(dest), "note": "copy file"})
        actions.append(
            f"{'[DRY]' if dry_run else '[OK] '} File  -> {dest}"
        )
        if not dry_run:
            shutil.copy2(src, dest)
        return actions, planned

    # ── Folder contents only ──────────────────────────────────────────────────
    if src.is_dir():
        actions.append(f"Source folder: {src}")

        if disallow_symlinks:
            bad = _find_symlink_in_tree(src)
            if bad:
                raise CopySourceError(f"Symlink found inside source: {bad}")

        items = list(src.iterdir())

        # Pre-flight: check all destinations before touching anything.
        for item in items:
            dest = target_dir / item.name
            if dest.exists():
                raise CopySourceError(f"Destination already exists: {dest}")

        copied: list[Path] = []
        try:
            for item in items:
                dest = target_dir / item.name
                if item.is_file():
                    planned.append({
                        "kind": "FILE", "path": str(dest),
                        "note": "from folder contents",
                    })
                    actions.append(
                        f"{'[DRY]' if dry_run else '[OK] '} File  -> {dest}"
                    )
                    if not dry_run:
                        shutil.copy2(item, dest)
                        copied.append(dest)
                else:
                    if dest.exists():
                        raise CopySourceError(f"Destination already exists: {dest}")
                    planned.append({
                        "kind": "DIR", "path": str(dest),
                        "note": "subfolder from contents",
                    })
                    actions.append(
                        f"{'[DRY]' if dry_run else '[OK] '} Dir   -> {dest}"
                    )
                    if not dry_run:
                        shutil.copytree(item, dest)
                        copied.append(dest)

        except (CopySourceError, Exception) as exc:
            if not dry_run and copied:
                actions.append(f"[ROLLBACK] Removing {len(copied)} item(s)…")
                for p in reversed(copied):
                    try:
                        if p.is_dir():
                            shutil.rmtree(p, ignore_errors=True)
                        elif p.exists():
                            p.unlink(missing_ok=True)
                    except Exception:
                        pass
            if isinstance(exc, CopySourceError):
                raise
            raise CopySourceError(str(exc)) from exc

        return actions, planned

    raise CopySourceError(f"Source is neither file nor directory: {src}")


def execute_job(
    job: Job,
    output_root: Path,
    *,
    allow_absolute: bool,
    forbid_traversal: bool,
    exist_ok: bool,
    disallow_symlinks: bool,
    dry_run: bool,
) -> tuple[Path, list[str], list[dict]]:
    actions: list[str] = []
    planned: list[dict] = []

    target = compute_target_path(
        output_root, job.structure_path,
        allow_absolute=allow_absolute,
        forbid_traversal=forbid_traversal,
    )

    planned.append({"kind": "DIR", "path": str(target), "note": "structure path"})
    ensure_dir(target, exist_ok=exist_ok, dry_run=dry_run)
    actions.append(
        f"[STRUCTURE] {'(dry) ' if dry_run else ''}Directory: {target}"
    )

    if job.copy_source and str(job.copy_source).strip():
        actions.append(f"[COPY] Source: {job.copy_source}")
        ca, cp = copy_into_target(
            job.copy_source, target,
            disallow_symlinks=disallow_symlinks,
            dry_run=dry_run,
        )
        actions.extend(ca)
        planned.extend(cp)
    else:
        actions.append("[COPY] No copy source for this job.")

    return target, actions, planned


# ==============================================================================
# Job JSON serialisation
# ==============================================================================
def jobs_to_json(jobs: list[Job]) -> str:
    data = [
        {
            "structure_path": j.structure_path,
            "copy_source": str(j.copy_source) if j.copy_source else None,
        }
        for j in jobs
    ]
    return json.dumps(data, ensure_ascii=False, indent=2)


def jobs_from_json(text: str) -> list[Job]:
    data = json.loads(text)
    result: list[Job] = []
    for item in data:
        cp = item.get("copy_source")
        result.append(Job(
            structure_path=item["structure_path"],
            copy_source=Path(cp) if cp else None,
        ))
    return result


# ==============================================================================
# Job Dialog
# ==============================================================================
class JobDialog(tk.Toplevel):
    def __init__(
        self, master: tk.Tk, title: str, initial: Job | None = None
    ) -> None:
        super().__init__(master)
        self.title(title)
        self.minsize(700, 220)
        self.resizable(True, False)
        self.transient(master)
        self.configure(bg=_C["bg"])

        self.result: Job | None = None

        self._sv = tk.StringVar(value=initial.structure_path if initial else "")
        self._cv = tk.StringVar(
            value=str(initial.copy_source) if (initial and initial.copy_source) else ""
        )

        pad = CONFIG["padding"]
        frm = ttk.Frame(self, padding=(pad, pad, pad, 8))
        frm.pack(fill=tk.BOTH, expand=True)

        # Structure Path
        ttk.Label(frm, text="Structure Path  (relative folder structure):").grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 4))
        e1 = ttk.Entry(frm, textvariable=self._sv)
        e1.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(0, 14))

        # Copy Source
        ttk.Label(frm, text="Copy Source  (optional — file or folder):").grid(
            row=2, column=0, columnspan=3, sticky="w", pady=(0, 4))
        e2 = ttk.Entry(frm, textvariable=self._cv)
        e2.grid(row=3, column=0, sticky="ew", padx=(0, 8))
        ttk.Button(frm, text="Pick File…",   command=self._pick_file).grid(
            row=3, column=1, padx=(0, 6))
        ttk.Button(frm, text="Pick Folder…", command=self._pick_folder).grid(
            row=3, column=2)

        # Hint
        ttk.Label(
            frm,
            text="If Copy Source is a folder — its CONTENTS are copied (no container).",
            style="Dim.TLabel",
        ).grid(row=4, column=0, columnspan=3, sticky="w", pady=(8, 0))

        frm.columnconfigure(0, weight=1)

        # Buttons
        btn_row = ttk.Frame(self, padding=(pad, 6, pad, pad))
        btn_row.pack(fill=tk.X)
        ttk.Button(btn_row, text="Cancel", command=self._cancel).pack(side=tk.RIGHT, padx=(8, 0))
        ttk.Button(btn_row, text="  Save  ", style="Accent.TButton",
                   command=self._save).pack(side=tk.RIGHT)

        # Focus / grab
        self.update_idletasks()
        self.lift()
        self.focus_force()
        e1.focus_set()
        self.grab_set()

        _bind_clipboard(e1)
        _bind_clipboard(e2)

        self.bind("<Return>", lambda _: self._save())
        self.bind("<Escape>", lambda _: self._cancel())

    def _pick_file(self) -> None:
        p = filedialog.askopenfilename(title="Select file to copy")
        if p:
            self._cv.set(p)

    def _pick_folder(self) -> None:
        p = filedialog.askdirectory(title="Select folder to copy")
        if p:
            self._cv.set(p)

    def _save(self) -> None:
        structure = self._sv.get().strip()
        copy_raw  = self._cv.get().strip()

        if not structure:
            messagebox.showerror("Validation Error", "Structure Path is required.",
                                 parent=self)
            return

        copy_path: Path | None = None
        if copy_raw:
            try:
                copy_path = safe_resolve_user_path(copy_raw, resolve=False)
            except ValueError as exc:
                messagebox.showerror("Validation Error", str(exc), parent=self)
                return
            if not copy_path.exists():
                if not messagebox.askyesno(
                    "Warning — Path Not Found",
                    f"Copy source does not exist yet:\n\n{copy_path}\n\n"
                    "Add the job anyway? It will be re-checked at runtime.",
                    parent=self,
                ):
                    return

        self.result = Job(structure_path=structure, copy_source=copy_path)
        self.destroy()

    def _cancel(self) -> None:
        self.result = None
        self.destroy()


# ==============================================================================
# Main Application
# ==============================================================================
class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(CONFIG["app_title"])
        self.geometry(CONFIG["window_size"])
        self.minsize(*CONFIG["min_size"])

        _apply_theme(self)

        self._output_var = tk.StringVar(value=str(Path.cwd()))
        d = CONFIG["defaults"]
        self._allow_abs_var       = tk.BooleanVar(value=d["allow_absolute_structure_path"])
        self._forbid_trav_var     = tk.BooleanVar(value=d["forbid_traversal"])
        self._dry_run_var         = tk.BooleanVar(value=d["dry_run"])
        self._disallow_sym_var    = tk.BooleanVar(value=d["disallow_symlinks"])
        self._help_lang_var       = tk.StringVar(value="en")

        self._jobs: list[Job] = []
        self._status_text = tk.StringVar(value="Ready")
        self._job_count_text = tk.StringVar(value="Jobs: 0")

        self._build_ui()

    # ──────────────────────────────────────────────────────────────────────────
    # UI construction
    # ──────────────────────────────────────────────────────────────────────────
    def _build_ui(self) -> None:
        self._build_header()

        scroll_canvas = tk.Canvas(self, bg=_C["bg"], highlightthickness=0)
        scrollbar = ttk.Scrollbar(self, orient="vertical",
                                  command=scroll_canvas.yview)
        scroll_canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self._main = ttk.Frame(scroll_canvas, padding=CONFIG["padding"])
        self._main_window = scroll_canvas.create_window(
            (0, 0), window=self._main, anchor="nw")

        def _on_configure(e):
            scroll_canvas.configure(scrollregion=scroll_canvas.bbox("all"))
        def _on_canvas_resize(e):
            scroll_canvas.itemconfig(self._main_window, width=e.width)

        self._main.bind("<Configure>", _on_configure)
        scroll_canvas.bind("<Configure>", _on_canvas_resize)
        scroll_canvas.bind_all("<MouseWheel>",
            lambda e: scroll_canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

        self._build_destination()
        self._build_jobs_section()
        self._build_settings()
        self._build_log_section()
        self._build_statusbar()

    def _build_header(self) -> None:
        hdr = tk.Canvas(self, height=68, bg=_C["hdr"], highlightthickness=0)
        hdr.pack(fill=tk.X, side=tk.TOP)

        # accent stripe
        hdr.create_rectangle(0, 0, 6000, 3, fill=_C["accent"], outline="")

        # icon block
        hdr.create_rectangle(14, 12, 50, 56, fill=_C["accent"],
                              outline=_C["accent_h"], width=1)
        hdr.create_text(32, 34, text="FS", fill="#fff",
                        font=("Segoe UI", 13, "bold"))

        # title
        hdr.create_text(62, 22, anchor="w",
                        text="Folder Structure Builder",
                        fill=_C["text"], font=("Segoe UI", 14, "bold"))
        hdr.create_text(62, 46, anchor="w",
                        text="Build folder hierarchies and copy files in batch",
                        fill=_C["text_dim"], font=("Segoe UI", 9))

        # version badge
        hdr.create_text(1060, 34, anchor="e",
                        text=CONFIG["version"],
                        fill=_C["accent"], font=("Segoe UI", 9, "bold"))

    def _build_destination(self) -> None:
        box = ttk.LabelFrame(self._main, text="  Destination", padding=10)
        box.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(box, text="Output Root  (base directory where folders will be created):",
                  style="Card.TLabel").pack(anchor="w", pady=(0, 6))

        row = ttk.Frame(box, style="Card.TFrame")
        row.pack(fill=tk.X)
        oe = ttk.Entry(row, textvariable=self._output_var)
        oe.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        _bind_clipboard(oe)
        b = ttk.Button(row, text="Browse…", command=self._browse_output)
        b.pack(side=tk.LEFT)
        _Tooltip(b, "Choose the root output folder")

    def _build_jobs_section(self) -> None:
        box = ttk.LabelFrame(self._main, text="  Jobs", padding=10)
        box.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # Treeview
        tf = ttk.Frame(box)
        tf.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        cols = ("idx", "structure", "copy")
        self._tree = ttk.Treeview(tf, columns=cols, show="headings", height=9)
        self._tree.heading("idx",       text="#",                anchor="center")
        self._tree.heading("structure", text="Structure Path",   anchor="w")
        self._tree.heading("copy",      text="Copy Source",      anchor="w")
        self._tree.column("idx",       width=46, anchor="center", stretch=False)
        self._tree.column("structure", width=430, stretch=True)
        self._tree.column("copy",      width=400, stretch=True)
        self._tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._tree.bind("<Double-1>", lambda _: self._edit_job())
        self._tree.tag_configure("alt", background=_C["surface2"])

        sb = ttk.Scrollbar(tf, orient="vertical", command=self._tree.yview)
        self._tree.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        # Action buttons
        br = ttk.Frame(box)
        br.pack(fill=tk.X)

        def btn(parent, text, cmd, style="TButton", tip=""):
            b = ttk.Button(parent, text=text, command=cmd, style=style)
            b.pack(side=tk.LEFT, padx=(0, 6))
            if tip:
                _Tooltip(b, tip)
            return b

        btn(br, "＋  Add",        self._add_job,       "Accent.TButton", "Add a new job (Ins)")
        btn(br, "✎  Edit",        self._edit_job,       tip="Edit selected job (F2)")
        btn(br, "⧉  Duplicate",   self._duplicate_job,  tip="Duplicate selected job")
        btn(br, "↑  Move Up",     lambda: self._move(-1), tip="Move job up")
        btn(br, "↓  Move Down",   lambda: self._move(1),  tip="Move job down")

        # Right side
        btn(br, "🗑  Remove",  self._remove_job,  "Danger.TButton",
            tip="Remove selected job (Del)").pack_forget()
        rb = ttk.Button(br, text="🗑  Remove", command=self._remove_job,
                        style="Danger.TButton")
        rb.pack(side=tk.RIGHT, padx=(0, 0))
        _Tooltip(rb, "Remove selected job")

        cb = ttk.Button(br, text="Clear All", command=self._clear_jobs)
        cb.pack(side=tk.RIGHT, padx=(0, 8))
        _Tooltip(cb, "Remove all jobs")

        ib = ttk.Button(br, text="⬆ Import", command=self._import_jobs)
        ib.pack(side=tk.RIGHT, padx=(0, 6))
        _Tooltip(ib, "Import jobs from a JSON file")

        eb = ttk.Button(br, text="⬇ Export", command=self._export_jobs)
        eb.pack(side=tk.RIGHT, padx=(0, 6))
        _Tooltip(eb, "Export jobs to a JSON file")

        # Keyboard shortcuts
        self.bind("<Insert>", lambda _: self._add_job())
        self.bind("<F2>",     lambda _: self._edit_job())
        self.bind("<Delete>", lambda _: self._remove_job())

    def _build_settings(self) -> None:
        box = ttk.LabelFrame(self._main, text="  Settings", padding=10)
        box.pack(fill=tk.X, pady=(0, 10))

        r1 = ttk.Frame(box)
        r1.pack(fill=tk.X)
        for text, var in [
            ("Dry Run  (no filesystem changes)",  self._dry_run_var),
            ("Allow absolute structure path",      self._allow_abs_var),
            ("Forbid '..' traversal",              self._forbid_trav_var),
            ("Disallow symlinks",                  self._disallow_sym_var),
        ]:
            ttk.Checkbutton(r1, text=text, variable=var).pack(
                side=tk.LEFT, padx=(0, 18))

        r2 = ttk.Frame(box)
        r2.pack(fill=tk.X, pady=(8, 0))
        ttk.Label(r2, text="Help language:").pack(side=tk.LEFT)
        ttk.Combobox(r2, textvariable=self._help_lang_var,
                     values=["en", "ar"], state="readonly", width=6
                     ).pack(side=tk.LEFT, padx=(8, 0))
        hb = ttk.Button(r2, text="? Settings Help", command=self._open_help)
        hb.pack(side=tk.LEFT, padx=(12, 0))
        _Tooltip(hb, "Open detailed settings explanation")

    def _build_log_section(self) -> None:
        box = ttk.LabelFrame(self._main, text="  Log", padding=10)
        box.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # Run buttons row
        br = ttk.Frame(box)
        br.pack(fill=tk.X, pady=(0, 8))

        rb = ttk.Button(br, text="▶  Run All Jobs",
                        command=self._run_all, style="Accent.TButton")
        rb.pack(side=tk.LEFT)
        _Tooltip(rb, "Execute all jobs")

        pb = ttk.Button(br, text="👁  Preview Plan",
                        command=self._preview_plan)
        pb.pack(side=tk.LEFT, padx=(8, 0))
        _Tooltip(pb, "Simulate without making changes")

        cb = ttk.Button(br, text="Clear Log", command=self._clear_log)
        cb.pack(side=tk.RIGHT)
        _Tooltip(cb, "Clear the log area")

        # Progress bar (hidden until run starts)
        self._progress = ttk.Progressbar(box, mode="determinate")
        self._progress.pack(fill=tk.X, pady=(0, 6))
        self._progress.pack_forget()

        # Log text widget
        log_frame = ttk.Frame(box)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self._log = tk.Text(
            log_frame, height=12, wrap="word",
            bg=_C["surface"], fg=_C["text"],
            insertbackground=_C["accent2"],
            selectbackground=_C["accent"],
            selectforeground="#ffffff",
            relief="flat", bd=0,
            font=("Consolas", 9),
            state="normal",
        )
        self._log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        lsb = ttk.Scrollbar(log_frame, orient="vertical",
                             command=self._log.yview)
        self._log.configure(yscrollcommand=lsb.set)
        lsb.pack(side=tk.RIGHT, fill=tk.Y)

        # Colour tags
        self._log.tag_configure("ERROR",   foreground=_C["error"])
        self._log.tag_configure("SUCCESS", foreground=_C["success"])
        self._log.tag_configure("WARNING", foreground=_C["warning"])
        self._log.tag_configure("INFO",    foreground=_C["accent2"])
        self._log.tag_configure("DIM",     foreground=_C["text_dim"])
        self._log.tag_configure("SECTION",
            foreground=_C["accent"], font=("Consolas", 9, "bold"))
        self._log.tag_configure("JOB",
            foreground=_C["text"], font=("Consolas", 9, "bold"))

        self._log_msg("Ready — add jobs then press Run All Jobs.", "INFO")

    def _build_statusbar(self) -> None:
        self._statusbar = tk.Frame(self, bg=_C["surface"], height=26)
        self._statusbar.pack(fill=tk.X, side=tk.BOTTOM)

        tk.Label(
            self._statusbar, textvariable=self._status_text,
            bg=_C["surface"], fg=_C["text_dim"],
            font=("Segoe UI", 9), padx=12,
        ).pack(side=tk.LEFT, fill=tk.Y)

        tk.Frame(self._statusbar, bg=_C["border"], width=1).pack(
            side=tk.LEFT, fill=tk.Y, pady=4)

        tk.Label(
            self._statusbar, textvariable=self._job_count_text,
            bg=_C["surface"], fg=_C["accent"],
            font=("Segoe UI", 9, "bold"), padx=12,
        ).pack(side=tk.LEFT, fill=tk.Y)

        tk.Frame(self._statusbar, bg=_C["border"], height=1).pack(
            fill=tk.X, side=tk.TOP)

    # ──────────────────────────────────────────────────────────────────────────
    # Helpers
    # ──────────────────────────────────────────────────────────────────────────
    def _set_status(self, text: str) -> None:
        self._status_text.set(text)

    def _log_msg(self, msg: str, tag: str = "") -> None:
        self._log.config(state="normal")
        self._log.insert(tk.END, msg + "\n", (tag,) if tag else ())
        self._log.see(tk.END)

    def _clear_log(self) -> None:
        self._log.config(state="normal")
        self._log.delete("1.0", tk.END)
        self._log_msg("Log cleared.", "DIM")

    def _refresh_tree(self) -> None:
        for item in self._tree.get_children():
            self._tree.delete(item)
        for i, job in enumerate(self._jobs, start=1):
            cp = str(job.copy_source) if job.copy_source else "—"
            tag = "alt" if i % 2 == 0 else ""
            self._tree.insert("", "end", values=(i, job.structure_path, cp), tags=(tag,))
        self._job_count_text.set(f"Jobs: {len(self._jobs)}")

    def _selected_idx(self) -> int | None:
        sel = self._tree.selection()
        if not sel:
            return None
        v = self._tree.item(sel[0], "values")
        return int(v[0]) - 1 if v else None

    def _browse_output(self) -> None:
        p = filedialog.askdirectory(title="Select Output Root")
        if p:
            self._output_var.set(p)
            self._set_status(f"Output Root: {p}")
            self._log_msg(f"[UI] Output Root set to: {p}", "DIM")

    # ──────────────────────────────────────────────────────────────────────────
    # Job operations
    # ──────────────────────────────────────────────────────────────────────────
    def _add_job(self) -> None:
        dlg = JobDialog(self, "Add Job")
        self.wait_window(dlg)
        if dlg.result:
            self._jobs.append(dlg.result)
            self._refresh_tree()
            self._log_msg(f"[JOBS] Added job #{len(self._jobs)}", "INFO")
            self._set_status(f"Added job #{len(self._jobs)}")

    def _edit_job(self) -> None:
        idx = self._selected_idx()
        if idx is None:
            messagebox.showinfo("Select a job", "Click a job row first.")
            return
        dlg = JobDialog(self, "Edit Job", initial=self._jobs[idx])
        self.wait_window(dlg)
        if dlg.result:
            self._jobs[idx] = dlg.result
            self._refresh_tree()
            self._log_msg(f"[JOBS] Updated job #{idx + 1}", "INFO")
            self._set_status(f"Updated job #{idx + 1}")

    def _duplicate_job(self) -> None:
        idx = self._selected_idx()
        if idx is None:
            messagebox.showinfo("Select a job", "Click a job row first.")
            return
        dup = self._jobs[idx]
        self._jobs.insert(idx + 1, dup)
        self._refresh_tree()
        self._log_msg(f"[JOBS] Duplicated job #{idx + 1} -> #{idx + 2}", "INFO")
        self._set_status(f"Duplicated job #{idx + 1}")

    def _remove_job(self) -> None:
        idx = self._selected_idx()
        if idx is None:
            messagebox.showinfo("Select a job", "Click a job row first.")
            return
        removed = self._jobs.pop(idx)
        self._refresh_tree()
        self._log_msg(f"[JOBS] Removed job #{idx + 1}: {removed.structure_path}", "WARNING")
        self._set_status(f"Removed job #{idx + 1}")

    def _move(self, delta: int) -> None:
        idx = self._selected_idx()
        if idx is None:
            return
        ni = idx + delta
        if ni < 0 or ni >= len(self._jobs):
            return
        self._jobs[idx], self._jobs[ni] = self._jobs[ni], self._jobs[idx]
        self._refresh_tree()
        children = self._tree.get_children()
        if children and 0 <= ni < len(children):
            self._tree.selection_set(children[ni])
            self._tree.focus(children[ni])
            self._tree.see(children[ni])

    def _clear_jobs(self) -> None:
        if not self._jobs:
            return
        if messagebox.askyesno("Confirm", "Remove all jobs?"):
            self._jobs.clear()
            self._refresh_tree()
            self._log_msg("[JOBS] All jobs cleared.", "WARNING")
            self._set_status("All jobs cleared.")

    def _export_jobs(self) -> None:
        if not self._jobs:
            messagebox.showinfo("Nothing to export", "Add jobs first.")
            return
        p = filedialog.asksaveasfilename(
            title="Export Jobs",
            defaultextension=".json",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if not p:
            return
        try:
            Path(p).write_text(jobs_to_json(self._jobs), encoding="utf-8")
            self._log_msg(f"[EXPORT] Saved {len(self._jobs)} job(s) to: {p}", "SUCCESS")
            self._set_status(f"Exported {len(self._jobs)} job(s).")
        except Exception as exc:
            messagebox.showerror("Export Error", str(exc))

    def _import_jobs(self) -> None:
        p = filedialog.askopenfilename(
            title="Import Jobs",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if not p:
            return
        try:
            text = Path(p).read_text(encoding="utf-8")
            loaded = jobs_from_json(text)
        except Exception as exc:
            messagebox.showerror("Import Error", f"Could not load file:\n{exc}")
            return
        if not loaded:
            messagebox.showinfo("Empty file", "The file contained no jobs.")
            return
        if self._jobs:
            choice = messagebox.askyesnocancel(
                "Import Jobs",
                f"Found {len(loaded)} job(s).\n\n"
                "Yes = replace existing jobs\n"
                "No  = append to existing jobs\n"
                "Cancel = do nothing",
            )
            if choice is None:
                return
            if choice:
                self._jobs = loaded
            else:
                self._jobs.extend(loaded)
        else:
            self._jobs = loaded
        self._refresh_tree()
        self._log_msg(f"[IMPORT] Loaded {len(loaded)} job(s) from: {p}", "SUCCESS")
        self._set_status(f"Imported {len(loaded)} job(s).")

    # ──────────────────────────────────────────────────────────────────────────
    # Settings / Help
    # ──────────────────────────────────────────────────────────────────────────
    def _open_help(self) -> None:
        lang = (self._help_lang_var.get() or "en").lower()
        if lang not in ("en", "ar"):
            lang = "en"

        win = tk.Toplevel(self)
        win.title("Settings Help" if lang == "en" else "شرح الإعدادات")
        win.geometry("800x520")
        win.minsize(700, 400)
        win.configure(bg=_C["bg"])
        win.transient(self)

        pad = CONFIG["padding"]
        frm = ttk.Frame(win, padding=pad)
        frm.pack(fill=tk.BOTH, expand=True)

        top = ttk.Frame(frm)
        top.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(top, text="Language:" if lang == "en" else "اللغة:").pack(side=tk.LEFT)
        cb = ttk.Combobox(top, values=["en", "ar"], state="readonly", width=6)
        cb.set(lang)
        cb.pack(side=tk.LEFT, padx=(8, 0))

        tf = ttk.Frame(frm)
        tf.pack(fill=tk.BOTH, expand=True)
        txt = tk.Text(tf, wrap="word",
                      bg=_C["surface"], fg=_C["text"],
                      font=("Segoe UI", 10), relief="flat", bd=0,
                      padx=10, pady=8)
        txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb2 = ttk.Scrollbar(tf, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=sb2.set)
        sb2.pack(side=tk.RIGHT, fill=tk.Y)

        def _load(lang_: str) -> None:
            lang_ = (lang_ or "en").lower()
            if lang_ not in ("en", "ar"):
                lang_ = "en"
            self._help_lang_var.set(lang_)
            txt.config(state="normal")
            txt.delete("1.0", tk.END)
            txt.insert(tk.END, SETTINGS_HELP_TEXT[lang_])
            txt.config(state="disabled")
            win.title("Settings Help" if lang_ == "en" else "شرح الإعدادات")

        cb.bind("<<ComboboxSelected>>", lambda _: _load(cb.get()))
        _load(lang)

        bf = ttk.Frame(frm)
        bf.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(bf,
                   text="Close" if lang == "en" else "إغلاق",
                   command=win.destroy).pack(side=tk.RIGHT)

    # ──────────────────────────────────────────────────────────────────────────
    # Plan window (shared by Preview and Dry Run)
    # ──────────────────────────────────────────────────────────────────────────
    def _open_plan_window(self, plan_items: list[dict], title: str = "Plan") -> None:
        win = tk.Toplevel(self)
        win.title(title)
        win.geometry("1000x580")
        win.minsize(860, 460)
        win.configure(bg=_C["bg"])
        win.transient(self)

        pad = CONFIG["padding"]
        frm = ttk.Frame(win, padding=pad)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text=title,
                  font=("Segoe UI", 13, "bold")).pack(anchor="w", pady=(0, 8))

        tf = ttk.Frame(frm)
        tf.pack(fill=tk.BOTH, expand=True)

        cols = ("job", "kind", "path", "note")
        tree = ttk.Treeview(tf, columns=cols, show="headings", height=18)
        tree.heading("job",  text="Job #")
        tree.heading("kind", text="Type")
        tree.heading("path", text="Full Path")
        tree.heading("note", text="Note")
        tree.column("job",  width=60,  anchor="center", stretch=False)
        tree.column("kind", width=80,  anchor="center", stretch=False)
        tree.column("path", width=640, stretch=True)
        tree.column("note", width=240, stretch=True)
        tree.tag_configure("ERROR", foreground=_C["error"])
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        sb2 = ttk.Scrollbar(tf, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb2.set)
        sb2.pack(side=tk.RIGHT, fill=tk.Y)

        for item in plan_items:
            tag = "ERROR" if item.get("kind") == "ERROR" else ""
            tree.insert("", "end",
                        values=(item["job_index"], item["kind"],
                                item["path"], item.get("note", "")),
                        tags=(tag,))

        bf = ttk.Frame(win, padding=(pad, 6, pad, pad))
        bf.pack(fill=tk.X)

        def _copy_path() -> None:
            sel = tree.selection()
            if not sel:
                return
            val = tree.item(sel[0], "values")
            if val and val[2]:
                win.clipboard_clear()
                win.clipboard_append(val[2])
                self._set_status("Path copied to clipboard.")

        ttk.Button(bf, text="Copy Selected Path", command=_copy_path).pack(side=tk.LEFT)
        ttk.Button(bf, text="Close", command=win.destroy).pack(side=tk.RIGHT)

    # ──────────────────────────────────────────────────────────────────────────
    # Settings collection
    # ──────────────────────────────────────────────────────────────────────────
    def _collect_settings(self) -> dict:
        return {
            "allow_absolute":   self._allow_abs_var.get(),
            "forbid_traversal": self._forbid_trav_var.get(),
            "exist_ok":         True,
            "disallow_symlinks":self._disallow_sym_var.get(),
            "dry_run":          self._dry_run_var.get(),
        }

    # ──────────────────────────────────────────────────────────────────────────
    # Preview Plan
    # ──────────────────────────────────────────────────────────────────────────
    def _preview_plan(self) -> None:
        try:
            output_root = safe_resolve_user_path(self._output_var.get(), must_exist=True)
        except ValueError as exc:
            messagebox.showerror("Error", str(exc))
            return
        if not output_root.is_dir():
            messagebox.showerror("Error", "Output Root must be an existing directory.")
            return
        if not self._jobs:
            messagebox.showinfo("No jobs", "Add at least one job first.")
            return

        s = self._collect_settings()
        self._log_msg("─" * 50, "DIM")
        self._log_msg("PREVIEW PLAN  (simulated dry run)", "SECTION")

        plan_items: list[dict] = []

        for i, job in enumerate(self._jobs, start=1):
            try:
                _, _, planned = execute_job(
                    job, output_root,
                    allow_absolute=s["allow_absolute"],
                    forbid_traversal=s["forbid_traversal"],
                    exist_ok=s["exist_ok"],
                    disallow_symlinks=s["disallow_symlinks"],
                    dry_run=True,
                )
                for p in planned:
                    plan_items.append({
                        "job_index": i, "kind": p["kind"],
                        "path": p["path"], "note": p.get("note", ""),
                    })
            except (StructurePathError, CopySourceError) as exc:
                self._log_msg(f"[Job #{i}] {exc}", "ERROR")
                plan_items.append({
                    "job_index": i, "kind": "ERROR",
                    "path": "", "note": str(exc),
                })

        self._log_msg("PREVIEW END", "SECTION")
        if plan_items:
            self._open_plan_window(plan_items, "Preview Plan  (no changes made)")

    # ──────────────────────────────────────────────────────────────────────────
    # Run All
    # ──────────────────────────────────────────────────────────────────────────
    def _run_all(self) -> None:
        try:
            output_root = safe_resolve_user_path(self._output_var.get(), must_exist=True)
        except ValueError as exc:
            messagebox.showerror("Error", str(exc))
            return
        if not output_root.is_dir():
            messagebox.showerror("Error", "Output Root must be an existing directory.")
            return
        if not self._jobs:
            messagebox.showinfo("No jobs", "Add at least one job first.")
            return

        s   = self._collect_settings()
        dry = s["dry_run"]

        # Show progress bar
        self._progress.configure(maximum=len(self._jobs), value=0)
        self._progress.pack(fill=tk.X, pady=(0, 6))
        self.update_idletasks()

        self._log_msg("=" * 50, "SECTION")
        self._log_msg(
            f"RUN START  |  Jobs: {len(self._jobs)}  |  Dry: {dry}",
            "SECTION",
        )
        self._set_status("Running…")

        failures: list[str] = []
        successes = 0
        plan_items: list[dict] = []

        for i, job in enumerate(self._jobs, start=1):
            self._log_msg(f"  Job #{i}: {job.structure_path}", "JOB")

            try:
                segs = normalize_structure_path(
                    job.structure_path,
                    allow_absolute=s["allow_absolute"],
                    forbid_traversal=s["forbid_traversal"],
                )
                self._log_msg(f"   Segments: {segs}", "DIM")

                target, actions, planned = execute_job(
                    job, output_root,
                    allow_absolute=s["allow_absolute"],
                    forbid_traversal=s["forbid_traversal"],
                    exist_ok=s["exist_ok"],
                    disallow_symlinks=s["disallow_symlinks"],
                    dry_run=dry,
                )

                for a in actions:
                    tag = "WARNING" if "ROLLBACK" in a else "DIM"
                    self._log_msg(f"   {a}", tag)
                self._log_msg(f"   ✓ Target: {target}", "SUCCESS")

                if dry:
                    for p in planned:
                        plan_items.append({
                            "job_index": i, "kind": p["kind"],
                            "path": p["path"], "note": p.get("note", ""),
                        })

                successes += 1

            except (StructurePathError, CopySourceError) as exc:
                failures.append(f"Job #{i}: {exc}")
                self._log_msg(f"   ✗ {exc}", "ERROR")
            except PermissionError:
                failures.append(f"Job #{i}: Permission denied.")
                self._log_msg("   ✗ Permission denied.", "ERROR")
            except Exception as exc:
                failures.append(f"Job #{i}: {exc}")
                self._log_msg(f"   ✗ {exc}", "ERROR")

            self._progress["value"] = i
            self.update_idletasks()

        self._log_msg(
            f"RUN END  |  Success: {successes}/{len(self._jobs)}",
            "SECTION",
        )
        self._log_msg("=" * 50, "SECTION")

        # Hide progress bar
        self._progress.pack_forget()

        if failures:
            for f in failures:
                self._log_msg(f"  ✗ {f}", "ERROR")
            self._set_status(f"Done with errors — {successes}/{len(self._jobs)} ok")
            messagebox.showwarning(
                "Completed with Errors",
                f"Success: {successes}/{len(self._jobs)}\nSee log for details.",
            )
        else:
            self._set_status(f"All {successes} job(s) completed successfully.")
            messagebox.showinfo(
                "Success",
                f"All {successes} job(s) completed!"
                + ("\n(Dry Run — no changes made.)" if dry else ""),
            )

        if dry and plan_items:
            self._open_plan_window(plan_items, "Dry Run Plan  (no changes made)")


# ==============================================================================
# Entry Point
# ==============================================================================
def main() -> int:
    try:
        from ctypes import windll  # type: ignore
        windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    App().mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())