"""Módulos para transformar y subir listas de placas a un HVR Hikvision."""
from datetime import datetime
import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from requests.auth import HTTPDigestAuth
import requests
import pandas as pd
import config

# --- CONFIGURACIÓN ---
# Las configuraciones sensibles están en config.py
# ----------------------

# Intentar usar Pillow (soporta JPEG). Si no está disponible,
# usaremos tkinter.PhotoImage como fallback
try:
    from PIL import Image, ImageTk
except ImportError:
    Image = None
    ImageTk = None


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
        msg = (
            "No se pudo leer el archivo Excel. Comprueba que está en un "
            "formato válido (.xls o .xlsx). "
        )
        raise ValueError(msg)

    # Estructura esperada del archivo de entrada
    # (ID, NoCasa, Placa, Allowed/Blocked, MesPagado,
    #  NoFecha, NoRecibo, Marbete)
    df_out = pd.DataFrame()
    df_out["No."] = range(1, len(df) + 1)
    df_out["Plate No."] = df.iloc[:, 2]  # Columna de la placa

    df_out["Group(0 BlockList, 1 AllowList)"] = (
        df.iloc[:, 3].apply(
            lambda x: 1
            if str(x).strip().lower() in ["allow", "1", "permitido", "si"]
            else 0
        )
    )

    df_out[
        "Effective Start Date (Format: YYYY-MM-DD, eg., 2017-12-07)"
    ] = "2000-01-01"
    df_out[
        "Effective End Date (Format: YYYY-MM-DD, eg., 2017-12-07)"
    ] = datetime.now().strftime("%Y-%m") + "-10"

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

    return resp


def ejecutar():
    """Ejecuta el flujo principal: selecciona archivo, transforma y sube."""
    ruta = filedialog.askopenfilename(
        title="Selecciona el archivo de placas",
        filetypes=[("Excel files", "*.xls *.xlsx")],
    )
    if not ruta:
        return

    try:
        ruta_transformada = transformar_excel(ruta)
        resp = subir_archivo(ruta_transformada)

        if resp.status_code in [200, 201, 204]:
            messagebox.showinfo(
                "Éxito",
                f"Archivo subido correctamente.\n{ruta_transformada}",
            )
        else:
            messagebox.showerror(
                "Error",
                (
                    "No se pudo subir.\n"
                    f"Código: {resp.status_code}\n"
                    f"Detalle: {resp.text}"
                ),
            )
    except (
        FileNotFoundError,
        OSError,
        requests.RequestException,
        pd.errors.EmptyDataError,
        ValueError,
    ) as e:
        # Errores esperables: fichero no existe, problemas IO, errores HTTP,
        # o datos de Excel mal formateados.
        messagebox.showerror("Error crítico", str(e))


# --- Interfaz mínima ---
app = tk.Tk()
app.title("Actualizador de placas - Villa Linda 3")
app.geometry("500x500")
app.configure(bg="#f0f0f0")

imgFondo = config.IMAGE_DIR
background_label = None
image = None

# Cargar imagen de fondo de forma robusta y con 50% de opacidad si es posible
try:
    if os.path.exists(imgFondo):
        if Image is not None and ImageTk is not None:
            # Usar Pillow para soportar JPG y manejar opacidad
            pil_original = Image.open(imgFondo).convert("RGBA")

            # Obtener tamaño actual de la ventana (tras asignar geometry)
            app.update_idletasks()
            w, h = app.winfo_width(), app.winfo_height()
            if w <= 1 or h <= 1:
                # Fallback a 500x500 si winfo aún no reporta tamaño válido
                w, h = 500, 500

            # Elegir filtro de remuestreo compatible con varias versiones
            resample = getattr(Image, "Resampling", None)
            if resample is not None:
                resample_filter = resample.LANCZOS
            else:
                resample_filter = getattr(Image, "LANCZOS", Image.BICUBIC)

            pil_img = pil_original.resize((w, h), resample_filter)

            # Aplicar 50% de opacidad: si la imagen tiene canal alpha,
            # reemplazarlo por un canal uniforme al 50%.
            alpha = pil_img.split()[-1]
            # Crear un canal alpha uniforme con valor 128 (50%)
            alpha_mask = Image.new("L", pil_img.size, 128)
            pil_img.putalpha(alpha_mask)

            image = ImageTk.PhotoImage(pil_img)

            background_label = tk.Label(app, image=image)
            background_label.image = image
            # Colocar como fondo usando place para que otros widgets queden
            # encima. Luego enviarlo al fondo con lower().
            background_label.place(relx=0, rely=0, relwidth=1, relheight=1)
            background_label.lower()

            # Redimensionar dinámicamente al cambiar tamaño de ventana
            def _on_resize(event):
                try:
                    new_w, new_h = event.width, event.height
                    if new_w < 1 or new_h < 1:
                        return
                    resized = pil_original.resize(
                        (new_w, new_h), resample_filter
                    )
                    resized.putalpha(
                        Image.new("L", (new_w, new_h), 128)
                    )
                    tkimg = ImageTk.PhotoImage(resized)
                    background_label.configure(image=tkimg)
                    background_label.image = tkimg
                except (tk.TclError, OSError, ValueError):
                    # No interrumpir la app por errores de resize
                    return

            app.bind("<Configure>", _on_resize)
        else:
            # Si no hay Pillow, PhotoImage solo soporta PNG/GIF.
            try:
                image = tk.PhotoImage(file=imgFondo)
            except tk.TclError:
                image = None

            if image is not None:
                background_label = tk.Label(app, image=image)
                background_label.image = image
                background_label.place(relx=0, rely=0, relwidth=1, relheight=1)
                background_label.lower()
            else:
                print(
                    "Aviso: no se pudo cargar la imagen de fondo."
                    " Para opacidad JPG instale Pillow.",
                    file=sys.stderr,
                )
    else:
        print(
            f"Aviso: archivo de imagen de fondo no encontrado: {imgFondo}",
            file=sys.stderr,
        )
except (tk.TclError, OSError) as e:
    # Errores de lectura/decodificación de imagen; informar y continuar.
    print(f"Error cargando imagen de fondo: {e}", file=sys.stderr)

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
