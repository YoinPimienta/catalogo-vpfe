import re

filepath = 'templates/base_busqueda.html'
with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# Replace col-md-4 title
content = re.sub(
    r'<h6 class="fw-bold mb-3" style="font-size: 0.85rem;">Dirección de Impuestos y Aduanas Nacionales</h6>',
    r'<h6 class="fw-bold mb-3" style="font-size: 0.85rem;">Portal de Simulación para Auditoría</h6>',
    content
)
# If encoding caused it to look different in the file:
content = re.sub(
    r'<h6 class="fw-bold mb-3" style="font-size: 0.85rem;">Direcci.*?n de Impuestos y Aduanas Nacionales</h6>',
    r'<h6 class="fw-bold mb-3" style="font-size: 0.85rem;">Portal de Simulación para Auditoría</h6>',
    content
)

content = re.sub(
    r'<p class="mb-1" style="font-size: 0.75rem;">.*?Sede principal.*?Nivel Central.*?</p>',
    r'<p class="mb-1" style="font-size: 0.75rem;">Este portal es una simulación para auditoría.</p>',
    content
)
content = re.sub(
    r'<p class="mb-1" style="font-size: 0.75rem;">Edificio San Agust.*?n</p>',
    r'<p class="mb-1" style="font-size: 0.75rem;">No está afiliado a la DIAN</p>',
    content
)
content = re.sub(
    r'<p class="mb-1" style="font-size: 0.75rem;">C.*?digo Postal: 111711</p>',
    r'<p class="mb-1" style="font-size: 0.75rem;">Ni a ninguna entidad gubernamental.</p>',
    content
)
content = re.sub(
    r'<p class="mb-3" style="font-size: 0.75rem;">Horario Contact Center.*?</p>',
    r'<p class="mb-3" style="font-size: 0.75rem;">Horario Ficticio: 24/7 para propósitos de prueba.</p>',
    content
)

# Replace col-md-5 contents
content = re.sub(
    r'<p class="mb-1" style="font-size: 0.7rem;">57 601 489 9000.*?Impuestos</p>',
    r'<p class="mb-1" style="font-size: 0.7rem;">555-0000 Línea de Auditoría</p>',
    content
)
content = re.sub(
    r'<p class="mb-1" style="font-size: 0.7rem;">57 601 794 8880.*?Aduanas</p>',
    r'<p class="mb-1" style="font-size: 0.7rem;">555-0001 Línea de Pruebas</p>',
    content
)
content = re.sub(
    r'<p class="mb-3" style="font-size: 0.7rem;">57 601 794 8883.*?Cobro</p>',
    r'<p class="mb-3" style="font-size: 0.7rem;">555-0002 Soporte Simulado</p>',
    content
)

content = re.sub(
    r'<a href="#" class="text-white text-decoration-none text-decoration-underline-hover">PQSR y Denuncias</a>',
    r'<a href="#" class="text-white text-decoration-none text-decoration-underline-hover">Términos Simulados</a>',
    content
)
content = re.sub(
    r'<a href="#" class="text-white text-decoration-none text-decoration-underline-hover">Puntos de contacto</a>',
    r'<a href="#" class="text-white text-decoration-none text-decoration-underline-hover">Contacto de Prueba</a>',
    content
)

# Fallback block replacement for the entire footer text area if regex above fails:
import re
new_footer_text = '''                <!-- Columna 2 -->
                <div class="col-md-4 mb-4 border-start border-white-50 ps-4">
                    <h6 class="fw-bold mb-3" style="font-size: 0.85rem;">Portal de Simulación para Auditoría</h6>
                    <p class="mb-1" style="font-size: 0.75rem;">Este portal es una simulación para auditoría.</p>
                    <p class="mb-1" style="font-size: 0.75rem;">No está afiliado a la DIAN</p>
                    <p class="mb-1" style="font-size: 0.75rem;">Ni a ninguna entidad gubernamental.</p>
                    <p class="mb-3" style="font-size: 0.75rem;">Horario Ficticio: 24/7 para propósitos de prueba.</p>
                    
                    <div class="d-flex gap-2">
                        <a href="#" class="text-white bg-white bg-opacity-25 rounded-circle d-flex align-items-center justify-content-center text-decoration-none" style="width: 25px; height: 25px;"><i class="bi bi-twitter"></i></a>
                        <a href="#" class="text-white bg-white bg-opacity-25 rounded-circle d-flex align-items-center justify-content-center text-decoration-none" style="width: 25px; height: 25px;"><i class="bi bi-youtube"></i></a>
                        <a href="#" class="text-white bg-white bg-opacity-25 rounded-circle d-flex align-items-center justify-content-center text-decoration-none" style="width: 25px; height: 25px;"><i class="bi bi-linkedin"></i></a>
                        <a href="#" class="text-white bg-white bg-opacity-25 rounded-circle d-flex align-items-center justify-content-center text-decoration-none" style="width: 25px; height: 25px;"><i class="bi bi-facebook"></i></a>
                        <a href="#" class="text-white bg-white bg-opacity-25 rounded-circle d-flex align-items-center justify-content-center text-decoration-none" style="width: 25px; height: 25px;"><i class="bi bi-instagram"></i></a>
                    </div>
                </div>

                <!-- Columna 3 -->
                <div class="col-md-5 mb-0 px-md-4 pb-2 border-start border-white-50">
                    <div class="d-flex align-items-center mb-0 mt-n1">
                        <i class="bi bi-telephone-fill me-2 fs-6"></i>
                        <span class="fw-bold" style="font-size: 0.85rem;">Contacto</span>
                    </div>
                    <h6 class="fw-bold mb-2 mt-3" style="font-size: 0.85rem;">Contact Center Ficticio</h6>
                    <p class="mb-1" style="font-size: 0.7rem;">555-0000 Línea de Auditoría</p>
                    <p class="mb-1" style="font-size: 0.7rem;">555-0001 Línea de Pruebas</p>
                    <p class="mb-3" style="font-size: 0.7rem;">555-0002 Soporte Simulado</p>
                    
                    <h6 class="fw-bold mb-2" style="font-size: 0.85rem;">Servicio en Línea Ficticio</h6>
                    <div style="font-size: 0.7rem; line-height: 1.6;">
                        <a href="#" class="text-white text-decoration-none text-decoration-underline-hover">Política de Simulación</a> | 
                        <a href="#" class="text-white text-decoration-none text-decoration-underline-hover">Términos de Auditoría</a><br>
                        <a href="#" class="text-white text-decoration-none text-decoration-underline-hover">Privacidad Simulada</a>
                    </div>
                </div>'''

pattern = r'<!-- Columna 2 -->.*?<!-- Columna 3 -->.*?</div>\s*</div>'
content = re.sub(pattern, new_footer_text + '\n            </div>', content, flags=re.DOTALL)

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)
print('done')
