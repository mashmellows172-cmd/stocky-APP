"""Panel de Stocky: ajustes a la izquierda, vista previa real a la derecha."""

from __future__ import annotations

import base64
import os
import shutil
import tempfile
import threading
import tkinter as tk
from pathlib import Path
from tkinter import font as tkfont
from tkinter import simpledialog, ttk

from .areas import area_name, file_stem, header_label
from .export import output_dir, write_area_xlsx
from .models import StockReport
from .pdf import PdfError, export_pdf, print_xlsx
from .presets import DEFAULT_CONFIG_PATH, Formato, load_config, save_config
from .preview import render_pages
from .theme import COLORS, apply as apply_theme
from .ui_state import DEFAULT_AREAS, initial_selection, resolve_font, status_text

DENSITIES = {"Compacta": "compacta", "Normal": "normal", "Amplia": "amplia"}
FILLS = {"Sin color": "ninguna", "Gris": "gris", "Verde": "verde"}
MARGINS = {"0,8 cm": 0.8, "1,0 cm": 1.0, "1,5 cm": 1.5, "2,0 cm": 2.0}
ORIENTATIONS = {"Vertical": "vertical", "Horizontal": "horizontal"}

PREVIEW_SCALE = 1.4


def _label(mapping: dict[str, object], value) -> str:
    for text, item in mapping.items():
        if item == value:
            return text
    return next(iter(mapping))


def _miles(number: int) -> str:
    return f"{number:,}".replace(",", ".")


class App:
    def __init__(self, report: StockReport, source: Path, root: tk.Tk) -> None:
        self.report = report
        self.source = Path(source)
        self.root = root
        self.config_path = self._config_path()
        self.cfg = load_config(self.config_path)
        self.fmt = Formato.from_dict(self.cfg.last_format.to_dict())

        self.selection = initial_selection(sorted(report.areas()), self.cfg.last_areas)
        areas = sorted(report.areas())
        self.current_area = self.selection[0] if self.selection else (areas[0] if areas else 0)

        self.temp_dir = Path(tempfile.mkdtemp(prefix="stocky_"))
        self._after: str | None = None
        self._job = 0
        self._photos: dict[int, list[tk.PhotoImage]] = {}
        self._pdfs: dict[int, Path] = {}
        self._canvases: dict[int, tk.Canvas] = {}
        self._busy = False

        self._configure_window()
        self.fonts = sorted(set(tkfont.families(root)))
        apply_theme(root)
        self._build()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self._schedule()

    # ------------------------------------------------------------------ setup

    @staticmethod
    def _config_path() -> Path:
        base = os.environ.get("APPDATA")
        return Path(base) / "Stocky" / "config.json" if base else DEFAULT_CONFIG_PATH

    def _configure_window(self) -> None:
        self.root.title(f"Stocky - {self.report.fecha or self.source.stem}")
        width = min(1180, self.root.winfo_screenwidth() - 40)
        height = min(760, self.root.winfo_screenheight() - 90)
        x = max(0, (self.root.winfo_screenwidth() - width) // 2)
        y = max(0, (self.root.winfo_screenheight() - height) // 2 - 20)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        self.root.minsize(940, 600)
        self.root.configure(bg=COLORS["bg"])

    def _build(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)

        self._build_top()
        body = ttk.Frame(self.root, padding=(14, 10, 14, 6))
        body.grid(row=1, column=0, sticky="nsew")
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)
        self._build_sidebar(body)
        self._build_notebook(body)
        self._build_bottom()

    def _build_top(self) -> None:
        top = ttk.Frame(self.root, style="Top.TFrame", padding=(16, 12))
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(0, weight=1)
        ttk.Label(top, text=self.source.name, style="H1.TLabel").grid(row=0, column=0, sticky="w")
        self.summary = ttk.Label(top, text=status_text(self.report, self.selection), style="Muted.TLabel")
        self.summary.grid(row=1, column=0, sticky="w", pady=(2, 0))
        ttk.Button(top, text="Abrir carpeta", style="Ghost.TButton", command=self._open_folder).grid(
            row=0, column=1, rowspan=2, sticky="e"
        )

    def _build_sidebar(self, body: ttk.Frame) -> None:
        holder = ttk.Frame(body, style="Sidebar.TFrame", width=264)
        holder.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        holder.grid_propagate(False)
        holder.rowconfigure(0, weight=1)
        holder.columnconfigure(0, weight=1)

        canvas = tk.Canvas(holder, bg=COLORS["panel"], highlightthickness=0, width=248)
        canvas.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(holder, orient="vertical", command=canvas.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        canvas.configure(yscrollcommand=scroll.set)

        inner = ttk.Frame(canvas, style="Sidebar.TFrame", padding=(14, 12))
        window = canvas.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda event: canvas.itemconfigure(window, width=event.width))
        self._wheel(canvas, holder)

        inner.columnconfigure(0, weight=1)
        inner.columnconfigure(1, weight=1)

        self.v_font = tk.StringVar(value=resolve_font(self.fmt.font_family, self.fonts))
        self.v_size = tk.IntVar(value=self.fmt.font_size)
        self.v_title = tk.IntVar(value=self.fmt.title_size)
        self.v_density = tk.StringVar(value=_label(DENSITIES, self.fmt.density))
        self.v_zebra = tk.BooleanVar(value=self.fmt.zebra)
        self.v_fill = tk.StringVar(value=_label(FILLS, self.fmt.header_fill))
        self.v_margin = tk.StringVar(value=_label(MARGINS, self.fmt.margin_cm))
        self.v_orientation = tk.StringVar(value=_label(ORIENTATIONS, self.fmt.orientation))
        self.v_wide = tk.BooleanVar(value=self.fmt.wide_desc)
        self.v_rubro = tk.BooleanVar(value=self.fmt.show_rubro)
        self.v_thousands = tk.BooleanVar(value=self.fmt.thousands)
        self.v_extra = tk.StringVar(value=self.fmt.extra_header)

        variable_widgets = [
            ttk.Combobox(inner, textvariable=self.v_font, values=self.fonts, state="readonly"),
            ttk.Spinbox(inner, from_=8, to=16, textvariable=self.v_size, width=6),
            ttk.Spinbox(inner, from_=8, to=24, textvariable=self.v_title, width=6),
            ttk.Combobox(inner, textvariable=self.v_density, values=list(DENSITIES), state="readonly"),
            ttk.Combobox(inner, textvariable=self.v_fill, values=list(FILLS), state="readonly"),
            ttk.Combobox(inner, textvariable=self.v_margin, values=list(MARGINS), state="readonly"),
            ttk.Combobox(inner, textvariable=self.v_orientation, values=list(ORIENTATIONS), state="readonly"),
            ttk.Entry(inner, textvariable=self.v_extra),
        ]
        checks = [
            ("Filas alternadas", self.v_zebra),
            ("Descripción ancha", self.v_wide),
            ("Mostrar rubro", self.v_rubro),
            ("Separador de miles", self.v_thousands),
        ]

        row = 0
        for text, widget in [
            ("Fuente", variable_widgets[0]),
            ("Tamaño del cuerpo", variable_widgets[1]),
            ("Tamaño del título", variable_widgets[2]),
            ("Densidad", variable_widgets[3]),
            ("Color del encabezado", variable_widgets[4]),
            ("Márgenes", variable_widgets[5]),
            ("Orientación", variable_widgets[6]),
            ("Encabezado extra", variable_widgets[7]),
        ]:
            row = self._field(inner, row, text, widget)

        for text, variable in checks:
            check = ttk.Checkbutton(inner, text=text, variable=variable, style="Area.TCheckbutton")
            check.grid(row=row, column=0, columnspan=2, sticky="w", pady=(6, 0))
            row += 1

        ttk.Separator(inner, orient="horizontal").grid(row=row, column=0, columnspan=2, sticky="ew", pady=(14, 6))
        row += 1
        ttk.Label(inner, text=f"Áreas ({len(self.report.areas())})", style="Section.TLabel").grid(
            row=row, column=0, columnspan=2, sticky="w"
        )
        row += 1

        self.area_vars: dict[int, tk.BooleanVar] = {}
        for area, count in self.report.areas().items():
            variable = tk.BooleanVar(value=area in self.selection)
            name = area_name(area) or "SIN NOMBRE"
            self.area_vars[area] = variable
            check = ttk.Checkbutton(
                inner,
                text=f"{area:02d} · {name[:18]}  ({_miles(count)})",
                variable=variable,
                style="Area.TCheckbutton",
                command=self._on_area_toggle,
            )
            check.grid(row=row, column=0, columnspan=2, sticky="w")
            row += 1

        row += 1
        ttk.Button(inner, text="Todas", style="Ghost.TButton", command=self._select_all).grid(
            row=row, column=0, sticky="ew", pady=(4, 2))
        ttk.Button(inner, text="Las de siempre", style="Ghost.TButton", command=self._select_default).grid(
            row=row, column=1, sticky="ew", pady=(4, 2), padx=(4, 0))
        row += 1
        ttk.Button(inner, text="Guardar como preset", style="Ghost.TButton", command=self._save_preset).grid(
            row=row, column=0, columnspan=2, sticky="ew", pady=(2, 2))
        row += 1
        self.preset_button = ttk.Button(inner, text="Presets ▾", style="Ghost.TButton", command=self._show_presets)
        self.preset_button.grid(row=row, column=0, sticky="ew", pady=(2, 2))
        ttk.Button(inner, text="Restablecer", style="Ghost.TButton", command=self._reset_format).grid(
            row=row, column=1, sticky="ew", pady=(2, 2), padx=(4, 0))
        row += 1

        for variable in (
            self.v_font,
            self.v_size,
            self.v_title,
            self.v_density,
            self.v_zebra,
            self.v_fill,
            self.v_margin,
            self.v_orientation,
            self.v_wide,
            self.v_rubro,
            self.v_thousands,
            self.v_extra,
        ):
            variable.trace_add("write", self._schedule)

    def _field(self, parent: ttk.Frame, row: int, text: str, widget: tk.Widget) -> int:
        ttk.Label(parent, text=text, style="Sidebar.TLabel").grid(
            row=row, column=0, columnspan=2, sticky="w", pady=(8, 2)
        )
        widget.grid(row=row + 1, column=0, columnspan=2, sticky="ew")
        return row + 2

    def _build_notebook(self, body: ttk.Frame) -> None:
        self.notebook = ttk.Notebook(body)
        self.notebook.grid(row=0, column=1, sticky="nsew")
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)
        self._rebuild_tabs()

    def _build_bottom(self) -> None:
        bottom = ttk.Frame(self.root, style="Bottom.TFrame", padding=(16, 10))
        bottom.grid(row=2, column=0, sticky="ew")
        bottom.columnconfigure(0, weight=1)
        self.status_render = ttk.Label(bottom, text="Listo", style="Bottom.TLabel")
        self.status_render.grid(row=0, column=0, sticky="w")
        ttk.Button(bottom, text="Abrir PDF", style="Ghost.TButton", command=self._open_pdf).grid(
            row=0, column=1, padx=(8, 0))
        ttk.Button(bottom, text="Imprimir", style="Ghost.TButton", command=self._print).grid(
            row=0, column=2, padx=(8, 0))
        ttk.Button(bottom, text="Generar todos", style="Accent.TButton", command=self._generate_all).grid(
            row=0, column=3, padx=(8, 0))

    # -------------------------------------------------------------- utilidades

    def _wheel(self, canvas: tk.Canvas, hover: tk.Widget) -> None:
        def scroll(event: tk.Event) -> None:
            canvas.yview_scroll(-int(event.delta / 120), "units")

        hover.bind("<Enter>", lambda _event: canvas.bind_all("<MouseWheel>", scroll))
        hover.bind("<Leave>", lambda _event: canvas.unbind_all("<MouseWheel>"))

    def _set_status(self, text: str, error: bool = False) -> None:
        self.status_render.configure(text=text, foreground=COLORS["danger"] if error else COLORS["muted"])

    def _sync_format(self) -> None:
        def as_int(variable: tk.IntVar, fallback: int) -> int:
            try:
                return int(variable.get())
            except (tk.TclError, ValueError):
                return fallback

        self.fmt.font_family = resolve_font(self.v_font.get(), self.fonts)
        self.fmt.font_size = max(8, min(16, as_int(self.v_size, self.fmt.font_size)))
        self.fmt.title_size = max(8, min(24, as_int(self.v_title, self.fmt.title_size)))
        self.fmt.density = DENSITIES.get(self.v_density.get(), "normal")
        self.fmt.zebra = bool(self.v_zebra.get())
        self.fmt.header_fill = FILLS.get(self.v_fill.get(), "gris")
        self.fmt.margin_cm = MARGINS.get(self.v_margin.get(), 1.5)
        self.fmt.orientation = ORIENTATIONS.get(self.v_orientation.get(), "vertical")
        self.fmt.wide_desc = bool(self.v_wide.get())
        self.fmt.show_rubro = bool(self.v_rubro.get())
        self.fmt.thousands = bool(self.v_thousands.get())
        self.fmt.extra_header = self.v_extra.get().strip()

    def _load_vars_from_format(self) -> None:
        self.v_font.set(resolve_font(self.fmt.font_family, self.fonts))
        self.v_size.set(self.fmt.font_size)
        self.v_title.set(self.fmt.title_size)
        self.v_density.set(_label(DENSITIES, self.fmt.density))
        self.v_zebra.set(self.fmt.zebra)
        self.v_fill.set(_label(FILLS, self.fmt.header_fill))
        self.v_margin.set(_label(MARGINS, self.fmt.margin_cm))
        self.v_orientation.set(_label(ORIENTATIONS, self.fmt.orientation))
        self.v_wide.set(self.fmt.wide_desc)
        self.v_rubro.set(self.fmt.show_rubro)
        self.v_thousands.set(self.fmt.thousands)
        self.v_extra.set(self.fmt.extra_header)

    # ------------------------------------------------------------------ áreas

    def _select_all(self) -> None:
        for variable in self.area_vars.values():
            variable.set(True)
        self._on_area_toggle()

    def _select_default(self) -> None:
        for area, variable in self.area_vars.items():
            variable.set(area in DEFAULT_AREAS)
        self._on_area_toggle()

    def _on_area_toggle(self) -> None:
        self.selection = [area for area, variable in self.area_vars.items() if variable.get()]
        if self.selection and self.current_area not in self.selection:
            self.current_area = self.selection[0]
        self.summary.configure(text=status_text(self.report, self.selection))
        self._rebuild_tabs()
        self._schedule()

    def _rebuild_tabs(self) -> None:
        for tab in self.notebook.tabs():
            self.notebook.forget(tab)
        self._canvases.clear()
        for area in self.selection:
            frame = ttk.Frame(self.notebook)
            canvas = tk.Canvas(frame, bg=COLORS["preview_bg"], highlightthickness=0)
            bar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
            canvas.configure(yscrollcommand=bar.set)
            canvas.grid(row=0, column=0, sticky="nsew")
            bar.grid(row=0, column=1, sticky="ns")
            frame.rowconfigure(0, weight=1)
            frame.columnconfigure(0, weight=1)
            self._wheel(canvas, canvas)
            name = area_name(area) or "SIN NOMBRE"
            self.notebook.add(frame, text=f"{area:02d} · {name[:12]}")
            self._canvases[area] = canvas
        if self.selection:
            index = self.selection.index(self.current_area) if self.current_area in self.selection else 0
            self.notebook.select(index)

    def _on_tab_changed(self, _event=None) -> None:
        if not self.selection:
            return
        try:
            index = self.notebook.index(self.notebook.select())
        except tk.TclError:
            return
        if index >= len(self.selection):
            return
        area = self.selection[index]
        if area != self.current_area:
            self.current_area = area
            self._schedule()

    # ------------------------------------------------------------- vista previa

    def _schedule(self, *_args) -> None:
        if self._after is not None:
            try:
                self.root.after_cancel(self._after)
            except tk.TclError:
                pass
        self._after = self.root.after(400, self._refresh_preview)

    def _refresh_preview(self) -> None:
        self._after = None
        self._sync_format()
        if not self.selection:
            self._set_status("Marca al menos un área para ver la vista previa")
            return
        if self.current_area not in self.selection:
            self.current_area = self.selection[0]
        area = self.current_area
        self._job += 1
        job = self._job
        fmt = Formato.from_dict(self.fmt.to_dict())
        canvas = self._canvases.get(area)
        available = canvas.winfo_width() if canvas is not None else 0
        scale = (available - 28) / 612 if available > 240 else PREVIEW_SCALE
        scale = max(0.9, min(1.6, scale))
        self._set_status(f"Actualizando el área {area:02d}…")
        threading.Thread(target=self._render_worker, args=(area, fmt, job, scale), daemon=True).start()

    def _render_worker(self, area: int, fmt: Formato, job: int, scale: float) -> None:
        try:
            folder = self.temp_dir / f"{job:04d}"
            xlsx = write_area_xlsx(self.report, area, fmt, folder)
            if xlsx is None:
                self._post(job, lambda: self._set_status(f"El área {area:02d} no tiene ítems con stock", error=True))
                return
            pdf = export_pdf(xlsx, xlsx.with_suffix(".pdf"))
            pages = render_pages(pdf, scale=scale)
            images = [base64.b64encode(page).decode("ascii") for page in pages]
            self._post(job, lambda: self._paint(area, pdf, images, job))
        except PdfError as error:
            message = f"No se pudo generar el PDF: {error}"
            self._post(job, lambda: self._set_status(message, error=True))
        except Exception as error:  # pragma: no cover - defensa
            message = f"No se pudo generar la vista previa: {error}"
            self._post(job, lambda: self._set_status(message, error=True))

    def _post(self, job: int, action) -> None:
        def run() -> None:
            if job != self._job:
                return
            action()

        try:
            self.root.after(0, run)
        except tk.TclError:
            pass

    def _paint(self, area: int, pdf: Path, images: list[str], job: int) -> None:
        canvas = self._canvases.get(area)
        if canvas is None:
            return
        canvas.delete("all")
        photos: list[tk.PhotoImage] = []
        y = 16
        width = 0
        for data in images:
            photo = tk.PhotoImage(data=data)
            x = max(12, (canvas.winfo_width() - photo.width()) // 2)
            canvas.create_image(x, y, image=photo, anchor="nw")
            photos.append(photo)
            width = max(width, photo.width())
            y += photo.height() + 18
        canvas.configure(scrollregion=(0, 0, max(width + 24, canvas.winfo_width()), max(y, 1)))
        canvas.yview_moveto(0)
        canvas.xview_moveto(0)
        self._photos[area] = photos
        self._pdfs[area] = pdf
        self._cleanup_temp(pdf.parent)
        pages = len(images)
        self._set_status(f"Área {area:02d} · {pages} página{'s' if pages != 1 else ''} · {header_label(area)}")

    def _cleanup_temp(self, keep: Path) -> None:
        keep_dirs = {Path(item).parent for item in self._pdfs.values()}
        keep_dirs.add(Path(keep))
        for folder in self.temp_dir.iterdir():
            if folder.is_dir() and folder not in keep_dirs:
                shutil.rmtree(folder, ignore_errors=True)

    # ------------------------------------------------------------------ preset

    def _save_preset(self) -> None:
        name = simpledialog.askstring("Guardar preset", "Nombre del preset:", parent=self.root)
        if not name or not name.strip():
            return
        self._sync_format()
        name = name.strip()
        self.cfg.presets[name] = Formato.from_dict(self.fmt.to_dict())
        self._save_config()
        self._set_status(f"Preset «{name}» guardado")

    def _show_presets(self) -> None:
        menu = tk.Menu(self.root, tearoff=0)
        if not self.cfg.presets:
            menu.add_command(label="Sin presets guardados", state="disabled")
        for name in sorted(self.cfg.presets):
            menu.add_command(label=name, command=lambda preset=name: self._apply_preset(preset))
        button = self.preset_button
        menu.tk_popup(button.winfo_rootx(), button.winfo_rooty() + button.winfo_height())

    def _apply_preset(self, name: str) -> None:
        preset = self.cfg.presets.get(name)
        if preset is None:
            return
        self.fmt = Formato.from_dict(preset.to_dict())
        self._load_vars_from_format()
        self._set_status(f"Preset «{name}» aplicado")
        self._schedule()

    def _reset_format(self) -> None:
        self.fmt = Formato()
        self._load_vars_from_format()
        self._set_status("Formato restablecido")
        self._schedule()

    # ----------------------------------------------------------------- acciones

    def _open_folder(self) -> None:
        folder = output_dir(self.report, self.source)
        try:
            folder.mkdir(parents=True, exist_ok=True)
            os.startfile(str(folder))
        except OSError as error:
            self._set_status(f"No se pudo abrir la carpeta: {error}", error=True)

    def _open_pdf(self) -> None:
        pdf = self._pdfs.get(self.current_area)
        if pdf and pdf.exists():
            os.startfile(str(pdf))
        else:
            self._set_status("Todavía no hay PDF de esta área", error=True)

    def _print(self) -> None:
        if not self.selection:
            self._set_status("Marca al menos un área para imprimir", error=True)
            return
        area = self.current_area
        self._sync_format()
        fmt = Formato.from_dict(self.fmt.to_dict())
        self._set_status(f"Preparando el área {area:02d} para imprimir…")
        threading.Thread(target=self._print_worker, args=(area, fmt), daemon=True).start()

    def _print_worker(self, area: int, fmt: Formato) -> None:
        folder = Path(tempfile.mkdtemp(prefix="stocky_print_"))
        try:
            xlsx = write_area_xlsx(self.report, area, fmt, folder)
            if xlsx is None:
                message = f"El área {area:02d} no tiene ítems con stock"
                self.root.after(0, lambda: self._set_status(message, error=True))
                return
            print_xlsx(xlsx)
            self.root.after(0, lambda: self._set_status(f"Área {area:02d} enviada a la impresora"))
        except PdfError as error:
            message = str(error)
            self.root.after(0, lambda: self._set_status(message, error=True))
        except Exception as error:  # pragma: no cover - defensa
            message = f"No se pudo imprimir: {error}"
            self.root.after(0, lambda: self._set_status(message, error=True))
        finally:
            shutil.rmtree(folder, ignore_errors=True)

    def _generate_all(self) -> None:
        self._sync_format()
        if not self.selection:
            self._set_status("Marca al menos un área", error=True)
            return
        if self._busy:
            return
        self._busy = True
        fmt = Formato.from_dict(self.fmt.to_dict())
        areas = list(self.selection)
        folder = output_dir(self.report, self.source)
        self._set_status(f"Generando {len(areas)} área(s) en {folder.name}…")
        threading.Thread(target=self._generate_worker, args=(areas, fmt, folder), daemon=True).start()

    def _generate_worker(self, areas: list[int], fmt: Formato, folder: Path) -> None:
        made = 0
        errors: list[str] = []
        for area in areas:
            try:
                xlsx = write_area_xlsx(self.report, area, fmt, folder)
                if xlsx is None:
                    continue
                export_pdf(xlsx, xlsx.with_suffix(".pdf"))
                made += 1
            except Exception as error:
                errors.append(f"{file_stem(area)}: {error}")
        self.root.after(0, lambda: self._after_generate(folder, made, errors))

    def _after_generate(self, folder: Path, made: int, errors: list[str]) -> None:
        self._busy = False
        if errors:
            self._set_status(f"Guardé {made} área(s) en {folder}. Avisos: " + " | ".join(errors), error=True)
        elif made:
            self._set_status(f"Guardé {made} área(s) en {folder}")
        else:
            self._set_status("Ningún área seleccionada tiene ítems con stock", error=True)
        if made:
            try:
                os.startfile(str(folder))
            except OSError:
                pass

    # ------------------------------------------------------------------- cierre

    def _save_config(self) -> None:
        try:
            save_config(self.cfg, self.config_path)
        except OSError as error:
            self._set_status(f"No se pudo guardar la configuración: {error}", error=True)

    def _on_close(self) -> None:
        self._sync_format()
        self.cfg.last_format = Formato.from_dict(self.fmt.to_dict())
        self.cfg.last_areas = list(self.selection)
        self.cfg.last_folder = str(output_dir(self.report, self.source))
        self._save_config()
        self.root.destroy()


def _enable_dpi_awareness() -> None:
    """Evita que Windows estire la ventana en pantallas con escala."""
    try:
        import ctypes

        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:  # pragma: no cover - Windows antiguo o sin shcore
        pass


def run(report: StockReport, source: Path) -> None:
    _enable_dpi_awareness()
    root = tk.Tk()
    App(report, source, root)
    root.mainloop()
