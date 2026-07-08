import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd

from main import REQUIRED_COLUMNS, generate_all


# ── Constants ─────────────────────────────────────────────────────────────────
BG_COLOR      = "#F5F6FA"
ACCENT_COLOR  = "#4A6CF7"
WHITE         = "#FFFFFF"
TEXT_BLACK    = "#1A1A2E"
TEXT_GRAY     = "#888888"
SUCCESS_GREEN = "#27AE60"
ERROR_RED     = "#E74C3C"
WARN_ORANGE   = "#F39C12"
BORDER_COLOR  = "#DDE1EE"
FONT_FAMILY   = "Segoe UI"


# ── Helper: rounded-looking button ───────────────────────────────────────────
def make_button(parent, text, command, bg=ACCENT_COLOR, fg=WHITE, padx=18, pady=7):
    """
    Crea un botón con estilo personalizado (color de fondo azul por defecto).

    Parámetros:
        parent: El widget de tkinter donde se colocará el botón.
        text (str): El texto que mostrará el botón.
        command (callable): La función que se ejecutará al hacer click.
        bg (str): Color de fondo del botón en formato hexadecimal.
        fg (str): Color del texto del botón en formato hexadecimal.
        padx (int): Relleno horizontal interno del botón.
        pady (int): Relleno vertical interno del botón.

    Retorna:
        tk.Button: El botón ya configurado y listo para ser posicionado.
    """
    btn = tk.Button(
        parent, text=text, command=command,
        bg=bg, fg=fg, font=(FONT_FAMILY, 10, "bold"),
        relief="flat", cursor="hand2",
        padx=padx, pady=pady,
        activebackground=ACCENT_COLOR, activeforeground=WHITE,
    )
    return btn


# ── Helper: section label ─────────────────────────────────────────────────────
def make_label(parent, text, color=TEXT_BLACK, size=10, bold=False):
    """
    Crea un Label de texto con fuente y color personalizados.

    Parámetros:
        parent: El widget de tkinter donde se colocará el label.
        text (str): El texto que mostrará el label.
        color (str): Color del texto en formato hexadecimal.
        size (int): Tamaño de la fuente en puntos.
        bold (bool): Si es True, el texto aparecerá en negritas.

    Retorna:
        tk.Label: El label configurado listo para ser posicionado.
    """
    weight = "bold" if bold else "normal"
    return tk.Label(parent, text=text, bg=BG_COLOR,
                    fg=color, font=(FONT_FAMILY, size, weight))


# ── Column validation ─────────────────────────────────────────────────────────
def validate_excel_columns(filepath):
    """
    Verifica que el archivo Excel seleccionado tenga exactamente las columnas requeridas
    en el orden correcto, tal como lo espera el programa para generar las evidencias.

    Lee únicamente la fila de encabezados del Excel (fila 2, índice 1) y compara los
    nombres de columna contra la lista REQUIRED_COLUMNS definida en main.py.

    Parámetros:
        filepath (str): Ruta completa al archivo .xlsx que se va a validar.

    Retorna:
        bool: True si el archivo tiene el formato correcto, False si no coincide.
    """
    try:
        df = pd.read_excel(filepath, sheet_name="TestCases", header=1, nrows=0)
        # Only keep non-unnamed columns for comparison
        actual_cols = [c for c in df.columns if not str(c).startswith("Unnamed")]
        return actual_cols == REQUIRED_COLUMNS
    except Exception:
        return False


# ── Progress window ───────────────────────────────────────────────────────────
class ProgressWindow(tk.Toplevel):
    """
    Ventana emergente que muestra una barra de progreso mientras se generan los documentos Word.

    Se abre al hacer click en 'Generar' y se actualiza en tiempo real conforme se va
    completando cada documento. Al terminar, se cierra sola y muestra el mensaje de éxito.

    Esta ventana corre la generación en un hilo separado (thread) para que la interfaz
    gráfica no se congele mientras Python trabaja.
    """

    def __init__(self, parent, excel_path, output_dir, on_complete):
        """
        Inicializa y muestra la ventana de progreso, y arranca el proceso de generación
        en un hilo de fondo para no bloquear la interfaz gráfica.

        Parámetros:
            parent: La ventana principal de tkinter que lanza esta ventana emergente.
            excel_path (str): Ruta del Excel seleccionado por el usuario.
            output_dir (str): Ruta de la carpeta destino para los documentos Word.
            on_complete (callable): Función que se llamará cuando la generación termine,
                                    recibiendo (total_generados, nombre_carpeta) como argumentos.
        """
        super().__init__(parent)
        self.title("Generando documentos...")
        self.resizable(False, False)
        self.configure(bg=BG_COLOR)
        self.grab_set()  # block interaction with main window

        self._on_complete  = on_complete
        self._excel_path   = excel_path
        self._output_dir   = output_dir
        self._error        = None

        # ── Layout ───────────────────────────────────────────────
        frame = tk.Frame(self, bg=BG_COLOR, padx=40, pady=30)
        frame.pack()

        make_label(frame, "Generando evidencias Word...", size=12, bold=True).pack(pady=(0, 16))

        self._progress_var = tk.DoubleVar(value=0)
        self._bar = ttk.Progressbar(
            frame, variable=self._progress_var,
            maximum=100, length=360, mode="determinate",
        )
        self._bar.pack(pady=(0, 10))

        self._status_lbl = make_label(frame, "Iniciando...", color=TEXT_GRAY, size=9)
        self._status_lbl.pack()

        # Center the window on screen
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width()  // 2) - (self.winfo_width()  // 2)
        y = parent.winfo_y() + (parent.winfo_height() // 2) - (self.winfo_height() // 2)
        self.geometry(f"+{x}+{y}")

        # Start generation in background thread
        threading.Thread(target=self._run, daemon=True).start()

    def _run(self):
        """
        Ejecuta la generación de documentos en un hilo de fondo.

        Llama a generate_all() de main.py pasándole un callback que actualiza
        la barra de progreso después de cada documento generado. Si ocurre un
        error en algún documento, lo captura y lo reporta en la interfaz.
        """
        try:
            total = generate_all(
                excel_path        = self._excel_path,
                output_dir        = self._output_dir,
                progress_callback = self._on_progress,
            )
            # Schedule UI update on the main thread
            self.after(300, lambda: self._finish(total))
        except Exception as exc:
            self._error = str(exc)
            self.after(0, self._show_error)

    def _on_progress(self, done, total):
        """
        Callback que recibe main.py después de guardar cada documento Word.

        Actualiza la barra de progreso y el texto de estado de forma segura
        desde el hilo de fondo usando self.after() para no violar las reglas de tkinter.

        Parámetros:
            done (int): Cantidad de documentos ya generados.
            total (int): Cantidad total de documentos a generar.
        """
        pct = (done / total) * 100
        self.after(0, lambda: self._progress_var.set(pct))
        self.after(0, lambda: self._status_lbl.config(
            text=f"Documento {done} de {total}  —  {pct:.0f}%"
        ))

    def _finish(self, total):
        """
        Se ejecuta cuando todos los documentos han sido generados correctamente.

        Cierra la ventana de progreso y llama al callback on_complete para que
        la ventana principal muestre el mensaje de éxito final.

        Parámetros:
            total (int): Número de documentos Word generados exitosamente.
        """
        folder_name = os.path.basename(self._output_dir)
        self.destroy()
        self._on_complete(total, folder_name)

    def _show_error(self):
        """
        Se ejecuta si ocurrió un error durante la generación de algún documento.

        Cierra la ventana de progreso y muestra un mensaje de error con el detalle
        del problema para que el usuario sepa qué falló y pueda tomar acción.
        """
        self.destroy()
        messagebox.showerror(
            "Error al generar",
            f"Ocurrió un error durante la generación de documentos:\n\n{self._error}"
        )


# ── Main application window ───────────────────────────────────────────────────
class App(tk.Tk):
    """
    Ventana principal de la aplicación GUI.

    Contiene todos los elementos visuales: selector de archivo Excel, selector de
    carpeta destino, mensajes de estado y el botón 'Generar'. Coordina la interacción
    del usuario y llama a main.py para ejecutar la generación de documentos.
    """

    def __init__(self):
        """
        Inicializa la ventana principal con todos sus componentes gráficos y
        las variables de estado internas que controlan qué archivo y carpeta
        ha seleccionado el usuario.
        """
        super().__init__()
        self.title("Generacion automatica de Evidencias Word")
        self.resizable(False, False)
        self.configure(bg=BG_COLOR)

        # Internal state
        self._excel_path  = None   # full path of the selected .xlsx file
        self._output_dir  = None   # full path of the selected output folder

        self._build_ui()

        # Center window on screen
        self.update_idletasks()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        ww, wh = self.winfo_width(), self.winfo_height()
        self.geometry(f"+{(sw - ww) // 2}+{(sh - wh) // 2}")

    # ── UI construction ───────────────────────────────────────────────────────
    def _build_ui(self):
        """
        Construye y posiciona todos los widgets de la ventana principal:
        encabezado, sección de selección de archivo, sección de carpeta destino,
        botón 'Generar' y el área de mensajes de estado en la parte inferior.
        """
        outer = tk.Frame(self, bg=BG_COLOR, padx=36, pady=28)
        outer.pack(fill="both", expand=True)

        # ── App header ────────────────────────────────────────────
        header = tk.Frame(outer, bg=BG_COLOR)
        header.pack(fill="x", pady=(0, 22))

        make_label(
            header,
            "Generación automática de Evidencias Word",
            size=14, bold=True,
        ).pack(side="left")

        # ── Divider ───────────────────────────────────────────────
        tk.Frame(outer, bg=BORDER_COLOR, height=1).pack(fill="x", pady=(0, 20))

        # ── Section 1: select Excel file ──────────────────────────
        self._build_file_section(outer)

        # ── Divider ───────────────────────────────────────────────
        tk.Frame(outer, bg=BORDER_COLOR, height=1).pack(fill="x", pady=(20, 20))

        # ── Section 2: select output folder ──────────────────────
        self._build_folder_section(outer)

        # ── Divider ───────────────────────────────────────────────
        tk.Frame(outer, bg=BORDER_COLOR, height=1).pack(fill="x", pady=(20, 20))

        # ── Section 3: Generate button ────────────────────────────
        btn_row = tk.Frame(outer, bg=BG_COLOR)
        btn_row.pack(fill="x")
        make_button(btn_row, "  ▶  Generar", self._on_generate).pack(side="right")

        # ── Status message area (bottom) ──────────────────────────
        self._status_frame = tk.Frame(outer, bg=BG_COLOR)
        self._status_frame.pack(fill="x", pady=(18, 0))

        self._status_lbl = tk.Label(
            self._status_frame, text="", bg=BG_COLOR,
            font=(FONT_FAMILY, 9), justify="center", wraplength=480,
        )
        self._status_lbl.pack()

    def _build_file_section(self, parent):
        """
        Construye el bloque de selección de archivo Excel con:
        - Label principal 'Selecciona tu matriz de casos de prueba'
        - Nota gris 'Only accept format .xlsx'
        - Botón 'Seleccionar Archivo' alineado a la derecha
        - Label que muestra el nombre del archivo seleccionado

        Parámetros:
            parent: El frame contenedor donde se agregarán los widgets.
        """
        row = tk.Frame(parent, bg=BG_COLOR)
        row.pack(fill="x")

        # Left: labels stacked
        left = tk.Frame(row, bg=BG_COLOR)
        left.pack(side="left", fill="x", expand=True)

        make_label(left, "Selecciona tu matriz de casos de prueba",
                   size=10, bold=True).pack(anchor="w")
        make_label(left, "Only accept format .xlsx",
                   color=TEXT_GRAY, size=8).pack(anchor="w", pady=(2, 0))

        # Right: button
        make_button(row, "Seleccionar Archivo", self._on_select_file).pack(side="right")

        # File name display below
        self._file_lbl = make_label(parent, "Ningún archivo seleccionado",
                                    color=TEXT_GRAY, size=8)
        self._file_lbl.pack(anchor="w", pady=(6, 0))

        # Clear file button
        clear_file_row = tk.Frame(parent, bg=BG_COLOR)
        clear_file_row.pack(fill="x", pady=(4, 0))
        make_button(
            clear_file_row, "🗑  Eliminar ruta de archivo",
            self._on_clear_file, bg=ERROR_RED,
        ).pack(side="right")

    def _build_folder_section(self, parent):
        """
        Construye el bloque de selección de carpeta destino con:
        - Label principal 'Seleccionar carpeta destino'
        - Descripción gris de su propósito
        - Botón 'Seleccionar Carpeta' alineado a la derecha
        - Label que muestra la ruta de la carpeta seleccionada

        Parámetros:
            parent: El frame contenedor donde se agregarán los widgets.
        """
        row = tk.Frame(parent, bg=BG_COLOR)
        row.pack(fill="x")

        left = tk.Frame(row, bg=BG_COLOR)
        left.pack(side="left", fill="x", expand=True)

        make_label(left, "Seleccionar carpeta destino",
                   size=10, bold=True).pack(anchor="w")
        make_label(left, "Carpeta donde se guardarán los documentos Word generados",
                   color=TEXT_GRAY, size=8).pack(anchor="w", pady=(2, 0))

        make_button(row, "Seleccionar Carpeta", self._on_select_folder).pack(side="right")

        self._folder_lbl = make_label(parent, "Ninguna carpeta seleccionada",
                                      color=TEXT_GRAY, size=8)
        self._folder_lbl.pack(anchor="w", pady=(6, 0))

        # Clear folder button
        clear_folder_row = tk.Frame(parent, bg=BG_COLOR)
        clear_folder_row.pack(fill="x", pady=(4, 0))
        make_button(
            clear_folder_row, "🗑  Eliminar ruta de carpeta",
            self._on_clear_folder, bg=ERROR_RED,
        ).pack(side="right")

    # ── Event handlers ────────────────────────────────────────────────────────
    def _on_select_file(self):
        """
        Maneja el click en el botón 'Seleccionar Archivo'.

        Abre el explorador de archivos del sistema operativo filtrado solo para .xlsx.
        Si el usuario selecciona un archivo con extensión incorrecta, muestra un error
        y vuelve a abrir el explorador para que lo intente de nuevo.
        Si el archivo tiene extensión .xlsx pero no cumple con el formato de columnas
        requerido, muestra un error de formato y descarta la selección.
        Si todo es correcto, actualiza el estado de la interfaz con el nombre del archivo
        y muestra el mensaje de éxito de validación.
        """
        while True:
            path = filedialog.askopenfilename(
                title="Selecciona la matriz de casos de prueba",
                filetypes=[("Archivos Excel", "*.xlsx")],
            )

            # User cancelled the dialog
            if not path:
                return

            # Wrong extension guard (in case OS bypasses the filter)
            if not path.lower().endswith(".xlsx"):
                self._show_inline_error(
                    "Formato no válido. Solo se aceptan archivos con extensión .xlsx.\n"
                    "Por favor selecciona el archivo correcto."
                )
                continue  # reopen the file dialog

            # Correct extension — now validate column structure
            if not validate_excel_columns(path):
                self._show_inline_error(
                    "El documento no cumple con el estándar de la matriz de casos de "
                    "prueba predefinida.\nPor favor contacté a QA Team.\n"
                    "Debe contener las siguientes columnas en el orden exacto:\n\n"
                    + "\n".join(REQUIRED_COLUMNS)
                )
                self._excel_path = None
                self._file_lbl.config(text="Ningún archivo seleccionado", fg=TEXT_GRAY)
                return

            # All good
            self._excel_path = path
            filename = os.path.basename(path)
            self._file_lbl.config(text=f"✔  {filename}", fg=SUCCESS_GREEN)
            self._set_status(
                "Archivo cumple con requisitos.\n"
                "Haz click en el botón 'Generar' para ver tus documentos Word.",
                color=SUCCESS_GREEN,
            )
            return

    def _on_select_folder(self):
        """
        Maneja el click en el botón 'Seleccionar Carpeta'.

        Abre el explorador de carpetas del sistema operativo para que el usuario
        elija dónde quiere que se guarden los documentos Word generados.
        Si selecciona una carpeta, actualiza el label con la ruta completa.
        """
        folder = filedialog.askdirectory(title="Selecciona la carpeta de destino")
        if folder:
            self._output_dir = folder
            self._folder_lbl.config(text=f"✔  {folder}", fg=SUCCESS_GREEN)

    def _on_clear_file(self):
        """
        Maneja el click en el botón 'Eliminar ruta de archivo'.

        Limpia la selección actual del archivo Excel, restablece el label de
        nombre de archivo a su estado inicial y borra el mensaje de estado
        de la parte inferior, dejando todo listo para una nueva selección.
        """
        self._excel_path = None
        self._file_lbl.config(text="Ningún archivo seleccionado", fg=TEXT_GRAY)
        self._set_status("")

    def _on_clear_folder(self):
        """
        Maneja el click en el botón 'Eliminar ruta de carpeta'.

        Limpia la selección actual de la carpeta destino y restablece el label
        de carpeta a su estado inicial, dejando todo listo para una nueva selección.
        """
        self._output_dir = None
        self._folder_lbl.config(text="Ninguna carpeta seleccionada", fg=TEXT_GRAY)

    def _on_generate(self):
        """
        Maneja el click en el botón 'Generar'.

        Valida que tanto el archivo Excel como la carpeta destino hayan sido
        seleccionados antes de iniciar. Si falta alguno, muestra un mensaje de
        advertencia claro para guiar al usuario. Si ambos están listos, abre la
        ventana de progreso y arranca la generación de documentos Word.
        """
        # Guard: no Excel selected
        if not self._excel_path:
            self._set_status(
                "⚠  Debes seleccionar tu matriz de casos de prueba (.xlsx) antes de generar.",
                color=WARN_ORANGE,
            )
            return

        # Guard: no output folder selected
        if not self._output_dir:
            self._set_status(
                "⚠  Debes seleccionar una carpeta destino donde se guardarán "
                "los documentos Word antes de continuar.",
                color=WARN_ORANGE,
            )
            return

        # Both ready — open progress window
        self._set_status("")
        ProgressWindow(
            parent      = self,
            excel_path  = self._excel_path,
            output_dir  = self._output_dir,
            on_complete = self._on_generation_complete,
        )

    def _on_generation_complete(self, total, folder_name):
        """
        Callback que se ejecuta cuando ProgressWindow termina de generar todos los documentos.

        Muestra una ventana emergente de éxito con el total de documentos generados
        y el nombre de la carpeta donde fueron guardados.

        Parámetros:
            total (int): Número de documentos Word generados exitosamente.
            folder_name (str): Nombre de la carpeta destino donde se guardaron.
        """
        messagebox.showinfo(
            "¡Generación completada!",
            f"Se han generado {total} documento(s) Word exitosamente.\n\n"
            f"Puedes visualizarlos en la carpeta: {folder_name}"
        )

    # ── Status helpers ────────────────────────────────────────────────────────
    def _set_status(self, text, color=TEXT_BLACK):
        """
        Actualiza el texto del área de mensajes de estado en la parte inferior de la ventana.

        Se usa para mostrar mensajes de validación, advertencias o confirmaciones
        de forma centralizada, cambiando el color según el tipo de mensaje.

        Parámetros:
            text (str): El mensaje que se mostrará. Si es cadena vacía, el área queda en blanco.
            color (str): Color del texto en formato hexadecimal.
                         Verde para éxito, naranja para advertencia, rojo para error.
        """
        self._status_lbl.config(text=text, fg=color)

    def _show_inline_error(self, message):
        """
        Muestra un diálogo de error modal con un botón 'Cerrar'.

        Se usa específicamente cuando el usuario selecciona un archivo con formato
        incorrecto, de modo que después de cerrar el error el explorador se vuelva
        a abrir para que corrija su selección.

        Parámetros:
            message (str): Texto del error que se mostrará en el cuadro de diálogo.
        """
        messagebox.showerror("Archivo no válido", message)


# ── Entry point ───────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app = App()
    app.mainloop()
