"""Colores y estilos del panel."""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk

COLORS = {
    "bg": "#F8FAFC",
    "panel": "#FFFFFF",
    "ink": "#0F172A",
    "muted": "#64748B",
    "line": "#E6E8EA",
    "slate": "#334155",
    "accent": "#059669",
    "accent_ink": "#FFFFFF",
    "preview_bg": "#EEF1F4",
    "danger": "#DC2626",
    "soft": "#F2F3F4",
}

FONT_STACK = ["Segoe UI", "Calibri", "Arial", "Tahoma", "Verdana", "Consolas"]


def apply(root: tk.Misc) -> None:
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:  # pragma: no cover - temas alternativos
        pass

    base = ("Segoe UI", 10)
    style.configure(".", font=base, background=COLORS["bg"], foreground=COLORS["ink"])

    style.configure("TFrame", background=COLORS["bg"])
    style.configure("Sidebar.TFrame", background=COLORS["panel"])
    style.configure("Top.TFrame", background=COLORS["panel"])
    style.configure("Bottom.TFrame", background=COLORS["panel"])

    style.configure("H1.TLabel", background=COLORS["panel"], foreground=COLORS["ink"], font=("Segoe UI Semibold", 16))
    style.configure("TLabel", background=COLORS["bg"], foreground=COLORS["ink"])
    style.configure("Sidebar.TLabel", background=COLORS["panel"], foreground=COLORS["ink"])
    style.configure(
        "Section.TLabel", background=COLORS["panel"], foreground=COLORS["slate"], font=("Segoe UI Semibold", 10)
    )
    style.configure("Muted.TLabel", background=COLORS["panel"], foreground=COLORS["muted"])
    style.configure("Bottom.TLabel", background=COLORS["panel"], foreground=COLORS["muted"])

    style.configure("Accent.TButton", background=COLORS["accent"], foreground=COLORS["accent_ink"], padding=(12, 6))
    style.map(
        "Accent.TButton",
        background=[("active", "#047857"), ("disabled", "#A7D7C5")],
        foreground=[("disabled", "#F8FAFC")],
    )
    style.configure("Ghost.TButton", background=COLORS["panel"], foreground=COLORS["slate"], padding=(10, 5))
    style.map("Ghost.TButton", background=[("active", COLORS["soft"])])

    style.configure("Area.TCheckbutton", background=COLORS["panel"], foreground=COLORS["ink"], padding=(2, 3))
    style.map("Area.TCheckbutton", background=[("active", COLORS["panel"])])

    style.configure("TNotebook", background=COLORS["bg"], borderwidth=0)
    style.configure("TNotebook.Tab", padding=(10, 6), font=("Segoe UI", 9))

    style.configure("TEntry", fieldbackground="#FFFFFF")
    style.configure("TCombobox", fieldbackground="#FFFFFF")
