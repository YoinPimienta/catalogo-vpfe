# Auditoria Project - Sistema de Gestión de Facturas DIAN

Sistema Django para la auditoría y gestión de facturas según estándares DIAN (Dirección de Impuestos y Aduanas Nacionales de Colombia).

## Requisitos

- Python 3.13+
- pip (gestor de paquetes de Python)

## Instalación Local

```bash
# Clonar el repositorio
git clone <tu-repositorio>
cd AuditoriaProject

# Crear entorno virtual (opcional pero recomendado)
python -m venv venv
source venv/Scripts/activate  # Windows
# source venv/bin/activate  # macOS/Linux

# Instalar dependencias
pip install -r requirements.txt

# Ejecutar migraciones
python manage.py migrate

# Crear superusuario (opcional)
python manage.py createsuperuser

# Ejecutar servidor de desarrollo
python manage.py runserver
```

Accede a: `http://localhost:8000`
Panel de admin: `http://localhost:8000/admin`

## Estructura del Proyecto

```
AuditoriaProject/
├── dian_sim/           # Configuración principal de Django
├── facturas/           # Aplicación principal
├── templates/          # Plantillas HTML
├── static/             # CSS e imágenes
├── media/              # Archivos subidos
├── db.sqlite3          # Base de datos (desarrollo)
├── manage.py           # Comandos de Django
├── requirements.txt    # Dependencias de Python
├── Procfile            # Configuración para Render
└── runtime.txt         # Versión de Python
```

## Despliegue en Render

Lee el archivo **[DESPLIEGUE_RENDER.md](DESPLIEGUE_RENDER.md)** para instrucciones detalladas.

Resumen rápido:
1. Sube el código a GitHub
2. Crea una nueva aplicación en Render
3. Configura las variables de entorno
4. ¡Listo! Tu app estará en línea

## Variables de Entorno

Copia `.env.example` a `.env` y configurar según sea necesario:

```bash
cp .env.example .env
```

## Configuración

- **Idioma:** Español (Colombia)
- **Zona horaria:** America/Bogota
- **Base de datos:** SQLite (desarrollo) / PostgreSQL (producción)
- **Archivo estático:** WhiteNoise

## Funcionalidades

- ✅ Gestión de facturas
- ✅ Panel de búsqueda avanzada
- ✅ Generación de QR para facturas
- ✅ Panel de administración Django

## Tecnologías

- Django 6.0.4
- Python 3.13
- gunicorn (servidor de producción)
- WhiteNoise (servir archivos estáticos)
- SQLite / PostgreSQL (base de datos)

## Desarrollo

Para desarrollar localmente:

```bash
# Crear migraciones después de cambios en models
python manage.py makemigrations

# Aplicar migraciones
python manage.py migrate

# Ejecutar tests
python manage.py test

# Verificar código
python manage.py check
```

## Licencia

Proyecto de auditoría DIAN - Uso interno.

## Soporte

Para problemas o preguntas sobre el despliegue en Render, consulta [DESPLIEGUE_RENDER.md](DESPLIEGUE_RENDER.md).
