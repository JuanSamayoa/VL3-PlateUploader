# VL3-PlateUploader

Aplicación para transformar archivos Excel de placas y subirlos a un HVR Hikvision.

## Instalación y Uso

1. Descarga el ejecutable `hikivision_api.exe` desde la carpeta `dist`.
2. Coloca el ejecutable en una carpeta junto con el archivo `config.py` (si deseas modificarlo).
3. Ejecuta `hikivision_api.exe`.

## Configuración

Los datos sensibles están en `config.py`. Edita este archivo para cambiar:

- `HVR_IP`: IP del HVR Hikvision.
- `USERNAME` y `PASSWORD`: Credenciales de autenticación.
- `IMAGE_DIR`: Ruta a la imagen de fondo.

Si modificas `config.py`, vuelve a empaquetar con PyInstaller si es necesario.

## Empaquetado

Para generar el ejecutable con PyInstaller:

```
pyinstaller --onefile --noconsole --add-data "config.py;." hikivision_api.py
```

Esto incluye `config.py` dentro del ejecutable.
