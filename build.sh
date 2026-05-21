#!/usr/bin/env bash
# exit on error
set -o errexit

pip install -r requirements.txt
python manage.py migrate --noinput
python manage.py shell < create_superuser.py
python manage.py collectstatic --noinput
