"""Módulos para transformar y subir listas de placas a un HVR Hikvision."""
from datetime import datetime
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from requests.auth import HTTPDigestAuth
import requests
import pandas as pd
import config
import threading
import re

# Definición de errores con códigos
ERROR_CODES = {
    "Err-001": "Formato de archivo inválido. El archivo Excel no se puede "
               "leer o no tiene el formato esperado.",
    "Err-002": "Problema de conexión. No se puede conectar al servidor "
               "Hikvision.",
    "Err-003": "Datos inválidos en el archivo. El archivo contiene datos "
               "que no cumplen con los requisitos.",
    "Err-004": "Placas inválidas. Algunas placas no tienen el formato "
               "correcto (ej. P123ABC).",
    "Err-005": "Error al procesar el archivo. Ocurrió un problema durante "
               "la transformación del archivo.",
    "Err-006": "Error al subir el archivo. No se pudo enviar el archivo "
               "al servidor.",
    "Err-007": "Error desconocido. Contacta al soporte técnico."
}


def transformar_excel(ruta_origen):
    """Transforma archivo Excel de placas a formato requerido por Hikvision."""
    # Leer Excel de forma robusta: probar engines openpyxl (xlsx) y xlrd (xls)
    _, ext = os.path.splitext(ruta_origen)
    ext = ext.lower()

    df = None
    attempts = []

    def try_engine(engine_name):
        """Intenta leer Excel con el engine especificado."""
        nonlocal df
        try:
            mes = datetime.now().month
            meses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
                     "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre",
                     "Diciembre"]
            mes_hoja = meses[mes - 1]
            df = pd.read_excel(
                ruta_origen,
                engine=engine_name,
                sheet_name=mes_hoja,
            )

            return True
        except ImportError as ie:
            attempts.append((engine_name, f"ImportError: {ie}"))
            return False
        except ValueError as ve:
            # pandas puede lanzar ValueError si el formato no coincide
            # con el engine seleccionado
            attempts.append((engine_name, f"ValueError: {ve}"))
            return False
        except (OSError, RuntimeError) as ex:
            attempts.append((engine_name, f"Error: {ex}"))
            return False

    # Orden heurístico según la extensión
    if ext == ".xlsx":
        if not try_engine("openpyxl"):
            try_engine("xlrd")
    elif ext == ".xls":
        if not try_engine("xlrd"):
            try_engine("openpyxl")
    else:
        # Desconocido: intentar ambos
        if not try_engine("openpyxl"):
            try_engine("xlrd")

    if df is None:
        raise ValueError("Err-001")

    # Procesar en chunks si el archivo es grande para evitar sobrecargar la PC
    chunk_size = 100 if len(df) > 500 else len(df)
    df_out = pd.DataFrame()
    for start in range(0, len(df), chunk_size):
        end = min(start + chunk_size, len(df))
        chunk = df.iloc[start:end]
        chunk_out = pd.DataFrame()
        chunk_out["No."] = range(start + 1, end + 1)
        chunk_out["Plate No."] = chunk.iloc[:, 2]  # Columna de la placa

        # Filtrar placas vacías
        chunk_out = chunk_out[
            chunk_out["Plate No."].notna() &
            (chunk_out["Plate No."].str.strip() != "")
        ]

        # Validar formato de placas en el chunk
        placa_pattern = re.compile(r'^[PACUO]\d{3}[A-Z]{3}$')
        invalid_plates = []
        for idx, placa in enumerate(chunk_out["Plate No."]):
            if not placa_pattern.match(str(placa).strip()):
                invalid_plates.append(f"Fila {start + idx + 1}: '{placa}'")
        if invalid_plates:
            raise ValueError("Err-004")

        # Manejar columna de activo/inactivo
        if len(chunk.columns) > 3:
            chunk_out["Group(0 BlockList, 1 AllowList)"] = (
                chunk.iloc[:, 3].apply(
                    lambda x: 1 if str(x).strip().lower() in
                    ["allow", "1", "permitido", "si"] else 0
                )
            )
        else:
            chunk_out["Group(0 BlockList, 1 AllowList)"] = 0

        chunk_out[
            "Effective Start Date (Format: YYYY-MM-DD, eg., 2017-12-07)"
        ] = "2000-01-01"
        chunk_out[
            "Effective End Date (Format: YYYY-MM-DD, eg., 2017-12-07)"
        ] = datetime.now().strftime("%Y-%m") + "-10"

        df_out = pd.concat([df_out, chunk_out], ignore_index=True)

    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    # Guardar como .xlsx usando openpyxl para evitar problemas con xlrd
    ruta_salida = os.path.join(
        os.path.dirname(ruta_origen),
        f"plateNolist_{config.HVR_IP}_{timestamp}.xlsx"
    )
    df_out.to_excel(ruta_salida, index=False, engine="openpyxl")
    return ruta_salida


def subir_archivo(ruta_archivo):
    """Sube archivo transformado al endpoint de Hikvision."""
    try:
        with open(ruta_archivo, 'rb') as f:
            files = {
                "file": (
                    os.path.basename(ruta_archivo),
                    f,
                    "application/octet-stream",
                )
            }
            resp = requests.put(
                config.UPLOAD_ENDPOINT,
                files=files,
                auth=HTTPDigestAuth(config.USERNAME, config.PASSWORD),
                timeout=15,
            )

        if resp.status_code not in [200, 201, 204]:
            raise Exception("Err-006")
        return resp
    except requests.exceptions.RequestException:
        raise Exception("Err-002")


def ejecutar():
    """Ejecuta el flujo principal: selecciona archivo, transforma y sube."""
    ruta = filedialog.askopenfilename(
        title="Selecciona el archivo de placas",
        filetypes=[("Excel files", "*.xls *.xlsx")],
    )
    if not ruta:
        return

    # Deshabilitar botón y mostrar loader
    btn.config(state="disabled")
    loader_label.config(text="Procesando archivo...")

    def proceso():
        try:
            ruta_transformada = transformar_excel(ruta)
            loader_label.config(text="Subiendo archivo...")
            resp = subir_archivo(ruta_transformada)

            if resp.status_code in [200, 201, 204]:
                messagebox.showinfo(
                    "Éxito",
                    "Archivo subido correctamente.",
                )
            else:
                messagebox.showerror(
                    "Error",
                    "No se pudo subir el archivo. Verifica la conexión.",
                )
        except ValueError as e:
            error_code = e.args[0] if e.args else "Err-003"
            msg = ERROR_CODES.get(error_code, "Error desconocido.")
            messagebox.showerror("Error", msg)
        except Exception as e:
            error_code = str(e) if str(e).startswith("Err-") else "Err-007"
            msg = ERROR_CODES.get(error_code, "Error desconocido.")
            messagebox.showerror("Error", msg)
        finally:
            # Rehabilitar botón y limpiar loader
            btn.config(state="normal")
            loader_label.config(text="")

    # Ejecutar en thread para no bloquear UI
    thread = threading.Thread(target=proceso)
    thread.start()


# --- Interfaz mínima ---
app = tk.Tk()
app.title("Actualizador de placas - Villa Linda 3")
app.geometry("500x500")
app.configure(bg="#f0f0f0")

# Loader label para indicar progreso
loader_label = tk.Label(app, text="", bg="#f0f0f0", font=("Arial", 12))
loader_label.pack(pady=10)

label = tk.Label(app,
                 text="Actualizador de placas - Villa Linda 3",
                 bg="#f0f0f0",
                 font=("Arial", 18))

label.pack(
    fill=tk.BOTH,
    padx=10,
    pady=10)

btn = tk.Label(app, text=("Seleccione para subir el archivo de placas con "
                          "extensión .xls o .xlsx"),
               font=("Arial", 18, "bold"),
               fg="#ffffff",
               bg="#007bff",
               wraplength=300,
               justify="center",
               relief="raised",
               bd=2)

btn.pack(
    fill=tk.BOTH,
    expand=True,
    padx=40,
    pady=40)

# Bind click event to execute function
btn.bind("<Button-1>", lambda e: ejecutar())

# Botón de salir pequeño en la parte inferior
btn_salir = tk.Button(
    app,
    text="Salir",
    command=app.quit,
    height=1,
    width=10,
    bg="#dc3545",
    fg="white"
)

btn_salir.pack(
    side=tk.BOTTOM,
    pady=10
)

app.mainloop()
