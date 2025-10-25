# VL3-PlateUploader

Aplicación GUI en Python para una colonia que transforma archivos Excel de placas de vehículos y los sube a un servidor HVR Hikvision.

## Características

- Interfaz gráfica simple con Tkinter.
- Lee archivos Excel (.xls/.xlsx) de la hoja del mes actual.
- Valida placas en formato: una letra (P/A/C/U/O), tres dígitos, tres letras mayúsculas (ej. P123ABC).
- Filtra automáticamente placas vacías.
- Asigna 0 por defecto si no hay columna de activo/inactivo (0=BlockList, 1=AllowList).
- Procesa archivos grandes en chunks para evitar sobrecarga.
- Subida asíncrona con threading para no bloquear la UI.
- Manejo de errores con códigos específicos (Err-001 a Err-007).
- Empaquetado como ejecutable portable (.exe) con PyInstaller.

## Requisitos

- Python 3.13.9
- Librerías: pandas, requests, tkinter, threading, re
- Archivo de configuración `config.py` con IP, usuario, contraseña del HVR.

## Instalación y Uso

1. Descarga el ejecutable `hikivision_api.exe` desde la carpeta `dist`.
2. Coloca el ejecutable junto con `config.py`.
3. Ejecuta `hikivision_api.exe`.
4. Selecciona un archivo Excel y haz clic para procesar y subir.

## Configuración

Edita `config.py` para:

- `HVR_IP`: Dirección IP del HVR Hikvision.
- `USERNAME` y `PASSWORD`: Credenciales.
- `UPLOAD_ENDPOINT`: URL de subida (basada en HVR_IP).

## Empaquetado

Genera el exe con:

```
pyinstaller --onefile --noconsole --hidden-import pandas --hidden-import requests hikivision_api.py
```

## Errores Comunes

- Err-001: Formato de archivo inválido.
- Err-002: Problema de conexión.
- Err-003: Datos inválidos.
- Err-004: Placas inválidas.
- Err-005: Error al procesar.
- Err-006: Error al subir.
- Err-007: Error desconocido.

---

Hecho por: Juan Samayoa <br>
Contacto por medio de web: [Portafolio - Sección Contacto](https://juan-samayoa.is-a.dev) o en [Mi Perfil](https://github.com/JuanSamayoa)
