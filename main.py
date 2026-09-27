import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# importar tus scripts
import ProcesarDPAH
import UnificarDPAH
import ProcesarCCE
import RealizarCruces


# =========================
# APP
# =========================
class App(tk.Tk):

    def __init__(self):
        super().__init__()

        self.title("Procesador Bancario")
        self.geometry("500x360")
        self.resizable(False, False)

        self._crear_interfaz()

    # =========================
    # UI
    # =========================
    def _crear_interfaz(self):

        frame = ttk.Frame(self, padding=20)
        frame.pack(fill="both", expand=True)

        ttk.Label(
            frame,
            text="Automatización de Procesos",
            font=("Segoe UI", 14, "bold")
        ).pack(pady=(0, 15))

        # BOTONES
        ttk.Button(frame, text="Procesar DPAH",
                   command=self.btn_dpah).pack(fill="x", pady=4)

        ttk.Button(frame, text="Unificar DPAH",
                   command=self.btn_unificar).pack(fill="x", pady=4)

        ttk.Button(frame, text="Procesar CCE",
                   command=self.btn_cce).pack(fill="x", pady=4)

        ttk.Button(frame, text="Realizar Cruces",
                   command=self.btn_cruces).pack(fill="x", pady=4)

        ttk.Separator(frame).pack(fill="x", pady=12)

        # PROGRESO
        self.progress = ttk.Progressbar(
            frame,
            mode="indeterminate"
        )
        self.progress.pack(fill="x", pady=5)

        # MENSAJES
        self.msg = tk.StringVar()
        ttk.Label(
            frame,
            textvariable=self.msg,
            wraplength=450
        ).pack(pady=10)

    # =========================
    # UTILIDADES
    # =========================
    def ejecutar_en_hilo(self, funcion, *args):

        def tarea():
            try:
                self.set_estado("Procesando...")
                self.progress.start()

                resultado = funcion(*args)

                self.set_estado(resultado or "Proceso completado ✓")

            except Exception as e:
                messagebox.showerror("Error", str(e))
                self.set_estado("Error durante el proceso")

            finally:
                self.progress.stop()

        threading.Thread(target=tarea, daemon=True).start()

    def set_estado(self, texto):
        self.msg.set(texto)

    # =========================
    # BOTONES
    # =========================
    def btn_dpah(self):
        archivos = filedialog.askopenfilenames(filetypes=[("TXT", "*.txt")])
        if archivos:
            self.ejecutar_en_hilo(ProcesarDPAH.ejecutar, archivos)

    def btn_unificar(self):
        archivos = filedialog.askopenfilenames(filetypes=[("Excel", "*.xlsx")])
        if archivos:
            self.ejecutar_en_hilo(UnificarDPAH.ejecutar, archivos)

    def btn_cce(self):
        archivo = filedialog.askopenfilename(filetypes=[("TXT", "*.txt")])
        if archivo:
            self.ejecutar_en_hilo(ProcesarCCE.ejecutar, archivo)

    def btn_cruces(self):
        cce = filedialog.askopenfilename(title="Archivo CCE", filetypes=[("Excel", "*.xlsx")])
        if not cce:
            return

        dpah = filedialog.askopenfilename(title="Archivo DPAH", filetypes=[("Excel", "*.xlsx")])
        if not dpah:
            return

        self.ejecutar_en_hilo(RealizarCruces.ejecutar, cce, dpah)


# =========================
# MAIN
# =========================
if __name__ == "__main__":
    App().mainloop()
