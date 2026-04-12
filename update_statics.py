import os
import re

template_dir = 'templates'
for filename in os.listdir(template_dir):
    if not filename.endswith('.html'): continue
    filepath = os.path.join(template_dir, filename)
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    # ensure {% load static %} is present
    if '{% load static %}' not in content:
        content = '{% load static %}\n' + content

    # replace /static/path/to/file with {% static 'path/to/file' %}
    content = re.sub(r'"/static/(.*?)"', r'"{% static \'\1\' %}"', content)

    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(content)
print("done")
