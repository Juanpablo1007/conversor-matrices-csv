import threading
from pathlib import Path
from tkinter import BOTH, END, LEFT, RIGHT, X, filedialog, messagebox
from tkinter import Tk
from tkinter import ttk

from excel_a_csv import (
    convert_excel_to_csv,
    convert_excel_to_extended_excel,
    convert_normal_csv_to_extended,
)


class ConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Conversor de matrices a CSV")
        self.root.geometry("820x520")
        self.root.minsize(720, 460)

        self.files = []
        self.output_folder = None
        self.is_running = False
        self.conversion_type = "csv_normal"

        self._configure_style()
        self._build_ui()

    def _configure_style(self):
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure("TFrame", background="#f6f7f9")
        style.configure("Header.TLabel", background="#f6f7f9", font=("Segoe UI", 16, "bold"))
        style.configure("Text.TLabel", background="#f6f7f9", font=("Segoe UI", 10))
        style.configure("Muted.TLabel", background="#f6f7f9", foreground="#5f6672", font=("Segoe UI", 9))
        style.configure("Success.TLabel", background="#f6f7f9", foreground="#147d3f", font=("Segoe UI", 10, "bold"))
        style.configure("Danger.TLabel", background="#f6f7f9", foreground="#a61b1b", font=("Segoe UI", 10, "bold"))
        style.configure("TButton", font=("Segoe UI", 10), padding=(12, 7))
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), padding=(14, 8))

    def _build_ui(self):
        main = ttk.Frame(self.root, padding=22)
        main.pack(fill=BOTH, expand=True)

        ttk.Label(main, text="Conversor de matrices de errores", style="Header.TLabel").pack(anchor="w")
        ttk.Label(
            main,
            text="Selecciona matrices Excel o CSV normales y genera el formato bancario requerido.",
            style="Muted.TLabel",
        ).pack(anchor="w", pady=(4, 18))

        actions = ttk.Frame(main)
        actions.pack(fill=X, pady=(0, 12))

        self.add_button = ttk.Button(actions, text="Agregar matrices", command=self.add_files)
        self.add_button.pack(side=LEFT)
        self.remove_button = ttk.Button(actions, text="Quitar seleccionada", command=self.remove_selected)
        self.remove_button.pack(side=LEFT, padx=(8, 0))
        self.clear_button = ttk.Button(actions, text="Limpiar lista", command=self.clear_files)
        self.clear_button.pack(side=LEFT, padx=(8, 0))

        self.convert_button = ttk.Button(
            actions,
            text="CSV normal",
            style="Accent.TButton",
            command=lambda: self.start_conversion("csv_normal"),
        )
        self.convert_button.pack(side=RIGHT)
        self.extended_button = ttk.Button(
            actions,
            text="CSV extendido",
            style="Accent.TButton",
            command=lambda: self.start_conversion("csv_extended"),
        )
        self.extended_button.pack(side=RIGHT, padx=(0, 8))
        self.extended_excel_button = ttk.Button(
            actions,
            text="Excel extendido",
            style="Accent.TButton",
            command=lambda: self.start_conversion("excel_extended"),
        )
        self.extended_excel_button.pack(side=RIGHT, padx=(0, 8))

        list_frame = ttk.Frame(main)
        list_frame.pack(fill=BOTH, expand=True)

        self.file_list = ttk.Treeview(
            list_frame,
            columns=("path", "status"),
            show="headings",
            selectmode="browse",
            height=10,
        )
        self.file_list.heading("path", text="Archivo de entrada")
        self.file_list.heading("status", text="Estado")
        self.file_list.column("path", width=560, minwidth=360)
        self.file_list.column("status", width=170, minwidth=130, anchor="center")
        self.file_list.pack(side=LEFT, fill=BOTH, expand=True)

        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.file_list.yview)
        scrollbar.pack(side=RIGHT, fill="y")
        self.file_list.configure(yscrollcommand=scrollbar.set)

        output_frame = ttk.Frame(main)
        output_frame.pack(fill=X, pady=(16, 8))

        self.output_button = ttk.Button(output_frame, text="Elegir carpeta de salida", command=self.choose_output_folder)
        self.output_button.pack(side=LEFT)
        self.output_label = ttk.Label(
            output_frame,
            text="Salida: junto a cada archivo",
            style="Text.TLabel",
        )
        self.output_label.pack(side=LEFT, padx=(12, 0), fill=X, expand=True)
        self.original_folder_button = ttk.Button(
            output_frame,
            text="Usar carpeta original",
            command=self.clear_output_folder,
        )
        self.original_folder_button.pack(side=RIGHT)

        self.progress = ttk.Progressbar(main, mode="determinate")
        self.progress.pack(fill=X, pady=(8, 8))

        self.status_label = ttk.Label(main, text="Listo para cargar matrices.", style="Muted.TLabel")
        self.status_label.pack(anchor="w")

    def add_files(self):
        if self.is_running:
            return

        selected = filedialog.askopenfilenames(
            title="Seleccionar matrices Excel o CSV normales",
            filetypes=[
                ("Matrices Excel y CSV", "*.xlsx *.xlsm *.csv"),
                ("Archivos Excel", "*.xlsx *.xlsm"),
                ("Archivos CSV", "*.csv"),
                ("Todos los archivos", "*.*"),
            ],
        )
        if not selected:
            return

        existing = {str(path) for path in self.files}
        for file_name in selected:
            file_path = Path(file_name)
            if str(file_path) in existing:
                continue
            self.files.append(file_path)
            self.file_list.insert("", END, values=(str(file_path), "Pendiente"))

        self.status_label.configure(text=f"Matrices cargadas: {len(self.files)}", style="Muted.TLabel")

    def remove_selected(self):
        if self.is_running:
            return

        selected = self.file_list.selection()
        if not selected:
            return

        item_id = selected[0]
        item_index = self.file_list.index(item_id)
        self.file_list.delete(item_id)
        del self.files[item_index]
        self.status_label.configure(text=f"Matrices cargadas: {len(self.files)}", style="Muted.TLabel")

    def clear_files(self):
        if self.is_running:
            return

        self.files.clear()
        for item_id in self.file_list.get_children():
            self.file_list.delete(item_id)
        self.progress.configure(value=0, maximum=100)
        self.status_label.configure(text="Lista limpia.", style="Muted.TLabel")

    def choose_output_folder(self):
        if self.is_running:
            return

        folder = filedialog.askdirectory(title="Seleccionar carpeta de salida")
        if not folder:
            return

        self.output_folder = Path(folder)
        self.output_label.configure(text=f"Salida: {self.output_folder}")

    def clear_output_folder(self):
        if self.is_running:
            return

        self.output_folder = None
        self.output_label.configure(text="Salida: junto a cada archivo")

    def start_conversion(self, conversion_type="csv_normal"):
        if self.is_running:
            return

        if not self.files:
            messagebox.showwarning("Sin archivos", "Agrega al menos una matriz Excel para convertir.")
            return

        self.is_running = True
        self.conversion_type = conversion_type
        self.convert_button.configure(state="disabled")
        self.extended_button.configure(state="disabled")
        self.extended_excel_button.configure(state="disabled")
        self.set_controls_state("disabled")
        self.progress.configure(value=0, maximum=len(self.files))
        if conversion_type == "excel_extended":
            self.status_label.configure(text="Generando Excel extendido...", style="Muted.TLabel")
        elif conversion_type == "csv_extended":
            self.status_label.configure(text="Generando version extendida...", style="Muted.TLabel")
        else:
            self.status_label.configure(text="Convirtiendo matrices...", style="Muted.TLabel")

        worker = threading.Thread(target=self.convert_files, daemon=True)
        worker.start()

    def convert_files(self):
        success_count = 0
        error_count = 0

        for index, input_path in enumerate(self.files):
            self.root.after(0, self.update_item_status, index, "Convirtiendo")

            try:
                output_path = self.build_output_path(input_path, self.conversion_type)
                if self.conversion_type == "excel_extended":
                    if input_path.suffix.lower() == ".csv":
                        raise RuntimeError(
                            "Los archivos CSV no se pueden convertir a Excel extendido."
                        )
                    output_path, row_count = convert_excel_to_extended_excel(
                        input_path,
                        output_path,
                    )
                elif input_path.suffix.lower() == ".csv":
                    if self.conversion_type != "csv_extended":
                        raise RuntimeError(
                            "Los archivos CSV solo se pueden convertir a CSV extendido."
                        )
                    output_path, row_count = convert_normal_csv_to_extended(
                        input_path,
                        output_path,
                    )
                else:
                    output_path, row_count = convert_excel_to_csv(
                        input_path,
                        output_path,
                        extended=self.conversion_type == "csv_extended",
                    )
                status = f"OK - {row_count} filas"
                success_count += 1
                self.root.after(0, self.update_item_status, index, status)
            except Exception as exc:
                error_count += 1
                self.root.after(0, self.update_item_status, index, f"Error: {exc}")

            self.root.after(0, self.progress.configure, {"value": index + 1})

        self.root.after(0, self.finish_conversion, success_count, error_count)

    def build_output_path(self, input_path, conversion_type):
        if self.output_folder is None and conversion_type == "csv_normal":
            return None

        if conversion_type == "excel_extended":
            file_name = f"{input_path.stem}_extendido.xlsx"
        elif conversion_type == "csv_extended":
            file_name = f"{input_path.stem}_extendido.csv"
        else:
            file_name = f"{input_path.stem}.csv"

        if self.output_folder is None:
            return input_path.with_name(file_name)
        return self.output_folder / file_name

    def update_item_status(self, index, status):
        item_id = self.file_list.get_children()[index]
        values = list(self.file_list.item(item_id, "values"))
        values[1] = status
        self.file_list.item(item_id, values=values)

    def finish_conversion(self, success_count, error_count):
        self.is_running = False
        self.convert_button.configure(state="normal")
        self.extended_button.configure(state="normal")
        self.extended_excel_button.configure(state="normal")
        self.set_controls_state("normal")

        if error_count:
            self.status_label.configure(
                text=f"Conversion terminada: {success_count} correctas, {error_count} con error.",
                style="Danger.TLabel",
            )
            messagebox.showwarning(
                "Conversion terminada",
                f"Se generaron {success_count} archivo(s) y hubo {error_count} error(es). Revisa la columna Estado.",
            )
        else:
            self.status_label.configure(
                text=f"Conversion terminada: {success_count} archivo(s) generado(s).",
                style="Success.TLabel",
            )
            messagebox.showinfo(
                "Conversion terminada",
                f"Se generaron {success_count} archivo(s) correctamente.",
            )

    def set_controls_state(self, state):
        for button in (
            self.add_button,
            self.remove_button,
            self.clear_button,
            self.output_button,
            self.original_folder_button,
        ):
            button.configure(state=state)


def main():
    root = Tk()
    ConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
