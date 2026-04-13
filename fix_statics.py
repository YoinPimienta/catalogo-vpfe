import os
import re

template_dir = 'templates'
modified_files = []

for filename in os.listdir(template_dir):
    if not filename.endswith('.html'): continue
    filepath = os.path.join(template_dir, filename)
    with open(filepath, 'r', encoding='utf-8') as f:
        original_content = f.read()

    # Fix the \' inside tags. It looks like {% static \'path\' %}
    content = re.sub(r"\{\%\s*static\s+\\\'(.+?)\\\'\s*\%\}", r"{% static '\1' %}", original_content)
    
    # Extract query params OUT of the static tag.
    # From: {% static 'path.css?v=8' %} to {% static 'path.css' %}?v=8
    # Match {% static 'SOMETHING?QUERY' %}
    content = re.sub(
        r"\{\%\s*static\s+'([^']+?)(\?[^']+)'\s*\%\}", 
        r"{% static '\1' %}\2", 
        content
    )

    if content != original_content:
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(content)
        modified_files.append(filename)

print("Modified files:", ", ".join(modified_files))
