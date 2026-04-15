web: gunicorn dian_sim.wsgi --log-file - --worker-class gthread --workers 1 --threads 4 --timeout 120
release: python manage.py migrate
