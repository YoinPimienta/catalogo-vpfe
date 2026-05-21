from django.contrib.auth import get_user_model
from decouple import config

User = get_user_model()

username = config("DJANGO_SUPERUSER_USERNAME", default=None)
email = config("DJANGO_SUPERUSER_EMAIL", default=None)
password = config("DJANGO_SUPERUSER_PASSWORD", default=None)

if username and email and password:
    if not User.objects.filter(username=username).exists():
        User.objects.create_superuser(
            username=username,
            email=email,
            password=password
        )
        print("Superuser created.")
    else:
        print("Superuser already exists.")
else:
    print("Superuser env vars not set, skipping creation.")
