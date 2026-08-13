# Cedulero

Sube acuses de pago del SAT en PDF y descarga la cédula de impuestos ya
llenada en Excel.

## Correr en local

```
pip install -r requirements.txt
python app.py
```

Abre `http://127.0.0.1:5000`.

## Desplegar en Render (para usarlo desde otra computadora)

1. Sube esta carpeta a un repositorio de GitHub (ver pasos abajo).
2. En [render.com](https://render.com), crea una cuenta y conecta tu GitHub.
3. **New +** → **Blueprint** → selecciona este repositorio. Render detecta
   `render.yaml` y configura el servicio solo.
   - Si prefieres configurarlo a mano: **New +** → **Web Service**, entorno
     *Python*, build command `pip install -r requirements.txt`, start
     command `gunicorn app:app --bind 0.0.0.0:$PORT --workers 1 --threads 4 --timeout 120`.
4. Al terminar el build, Render te da una URL pública
   (`https://cedulero-xxxx.onrender.com`). Ábrela desde cualquier
   computadora.

### Subir el código a GitHub

```
git init
git add app.py procesador.py index.html requirements.txt Procfile render.yaml .gitignore README.md plantilla
git commit -m "Cedulero: interfaz y backend"
git branch -M main
git remote add origin https://github.com/<tu-usuario>/cedulero.git
git push -u origin main
```

(Crea antes el repositorio vacío en github.com, o con `gh repo create`.)

## Limitaciones a tener en cuenta

- **Sin contraseña**: cualquiera con la URL puede subir PDFs y descargar las
  cédulas generadas. Si en algún momento quieres restringir el acceso,
  se puede agregar una contraseña simple.
- **Los archivos generados viven en memoria**, no en disco. Si el servicio
  se reinicia o se "duerme" (el plan gratuito de Render duerme el servicio
  tras ~15 minutos sin uso y tarda ~30-60s en despertar en la siguiente
  visita), se pierden las cédulas que no hayas descargado — tendrás que
  volver a subir los PDF.
- El `Procfile` usa un solo worker a propósito: si se aumentan los workers,
  las descargas pueden fallar porque cada worker tiene su propia memoria.
