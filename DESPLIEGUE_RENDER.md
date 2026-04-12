# Guía de Despliegue en Render

## Paso 1: Preparación Local (Ya completado ✓)

```bash
pip install gunicorn dj-database-url whitenoise
pip freeze > requirements.txt
```

Los archivos necesarios ya están configurados:
- ✅ **requirements.txt** - Dependencias del proyecto
- ✅ **Procfile** - Configuración para iniciar gunicorn
- ✅ **runtime.txt** - Versión de Python
- ✅ **settings.py** - Configurado para SQLite en local y PostgreSQL en Render
- ✅ **.gitignore** - Archivos a no subir
- ✅ **.env.example** - Variables de entorno

## Paso 2: Verificar Que Todo Funciona Localmente

```bash
# Ejecutar migraciones
python manage.py migrate

# Hacer migraciones si es necesario (opcional)
python manage.py makemigrations

# Crear un superusuario (opcional)
python manage.py createsuperuser

# Probar el servidor localmente
python manage.py runserver
```

Visita http://localhost:8000 para verificar que funciona.

## Paso 3: Subir a GitHub

```bash
# Iniciar repositorio git
git init

# Agregar archivos
git add .

# Commit inicial
git commit -m "Initial commit - Django project configured for Render"

# Agregar repositorio remoto (reemplaza con tu URL)
git remote add origin https://github.com/tu-usuario/tu-repo.git

# Subir a GitHub
git branch -M main
git push -u origin main
```

## Paso 4: Configurar en Render

1. **Accede a [render.com](https://render.com)** y crea una cuenta
2. **Crea un nuevo Web Service:**
   - Click en "New +" → "Web Service"
   - Conecta tu repositorio de GitHub
   - Selecciona el repositorio
   - Nombre: `dian-auditoria` (o el nombre que quieras)
   - Environment: `Python 3`
   - Build Command: `pip install -r requirements.txt`
   - Start Command: `gunicorn dian_sim.wsgi`

## Paso 5: Configurar Variables de Entorno en Render

En la sección "Environment" de tu servicio en Render, agrega estas variables:

```
DEBUG=False
SECRET_KEY=tu-clave-secreta-nueva-muy-larga
ALLOWED_HOSTS=tu-app.onrender.com,localhost
CSRF_TRUSTED_ORIGINS=https://tu-app.onrender.com
DATABASE_URL=postgresql://...  (Render proporciona esto si usas PostgreSQL)
```

## Paso 6: Crear Base de Datos (Opcional - Usar PostgreSQL de Render)

Si quieres usar PostgreSQL en lugar de SQLite en Render:

1. **Crea un PostgreSQL en Render:**
   - Click en "New +" → "PostgreSQL"
   - Nombre: `dian-auditoria-db`
   - Render te proporciona automáticamente la URL

2. **La variable DATABASE_URL se configurará automáticamente**

3. **Ejecuta las migraciones:**
   ```bash
   # En tu terminal local, conecta con la BD de Render:
   python manage.py migrate --settings=dian_sim.settings
   ```

## Paso 7: Desplegar

1. Render desplegará automáticamente cuando hagas push a GitHub
2. Puedes ver el estado del despliegue en el dashboard de Render
3. Una vez completado, tu app estará en: `https://tu-app.onrender.com`

## Solución de Problemas

### Erro: "Static files not found"
✅ Ya configurado con WhiteNoise en settings.py

### Error: "Connection to database refused"
- Asegúrate de que DATABASE_URL está correcta en Render
- Ejecuta migraciones en Render

### Error: "500 Internal Server Error"
- Revisa los logs en Render: Dashboard → tu app → "Logs"
- Verifica que DEBUG=False y que todas las variables de entorno están configuradas

### La aplicación se apaga después de 30 segundos
- Asegúrate de que el Procfile está correcto
- Verifica que gunicorn está instalado en requirements.txt

## Notas Importantes

- **No subas el archivo .env** con datos reales (ya está en .gitignore)
- **SECRET_KEY nueva:** Reemplaza la SECRET_KEY en Render con una nueva y segura
- **DEBUG siempre en False** en producción
- **Archivos media:** Si necesitas persistencia, considera usar un servicio como AWS S3 o Cloudinary

## Verificación Final

Después del despliegue:
1. ✅ Visita tu dominio en Render
2. ✅ Accede al panel de admin: `/admin`
3. ✅ Verifica que los estilos CSS se cargan correctamente
4. ✅ Comprueba que las imágenes se muestran
5. ✅ Prueba todas las funcionalidades de tu app
