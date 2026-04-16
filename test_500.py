import os
import django

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'dian_sim.settings')
django.setup()

from django.test import Client

c = Client()
response = c.post('/', {'cufe': '4b10f64a31ae4c9bae0cbf4905f4ce'})
print("Status Code:", response.status_code)
if response.status_code == 500:
    import traceback
    with open('error500.html', 'wb') as f:
        f.write(response.content)
    print("Error saved to error500.html")
