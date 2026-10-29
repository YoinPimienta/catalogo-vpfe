document.addEventListener('DOMContentLoaded', function() {
    const emisorNombreInput = document.getElementById('id_emisor_nombre');
    const plantillaInput = document.getElementById('id_plantilla');
    const serieInput = document.getElementById('id_serie');

    function sincronizarPlantillaSegunNombre(nombre) {
        if (!nombre || !plantillaInput) return;
        const norm = nombre.toUpperCase();
        if (norm.includes('EXITO') || norm.includes('ÉXITO')) {
            plantillaInput.value = 'EXITO';
            if (serieInput && (!serieInput.value || serieInput.value === 'AU06' || serieInput.value === 'I041')) {
                serieInput.value = 'WJ53';
            }
        } else if (norm.includes('CENCOSUD') || norm.includes('JUMBO')) {
            plantillaInput.value = 'JUMBO';
            if (serieInput && (!serieInput.value || serieInput.value === 'WJ53' || serieInput.value === 'I041')) {
                serieInput.value = 'AU06';
            }
        }
    }

    if (emisorNombreInput) {
        const loadingSpan = document.createElement('span');
        loadingSpan.id = 'emisor-loading';
        loadingSpan.style.display = 'none';
        loadingSpan.style.marginLeft = '10px';
        loadingSpan.style.color = '#94a3b8';
        loadingSpan.style.fontSize = '0.8rem';
        loadingSpan.textContent = 'Buscando datos...';
        emisorNombreInput.parentNode.appendChild(loadingSpan);

        emisorNombreInput.addEventListener('input', function() {
            sincronizarPlantillaSegunNombre(this.value);
        });

        emisorNombreInput.addEventListener('blur', function() {
            const nombre = this.value.trim();
            if (nombre) {
                sincronizarPlantillaSegunNombre(nombre);
                loadingSpan.style.display = 'inline';
                fetch(`/api/emisor-data/?emisor_nombre=${encodeURIComponent(nombre)}`)
                    .then(response => response.json())
                    .then(data => {
                        loadingSpan.style.display = 'none';
                        if (Object.keys(data).length > 0) {
                            if (data.plantilla && plantillaInput) {
                                plantillaInput.value = data.plantilla;
                            }
                            if (!document.getElementById('id_emisor_nit').value) {
                                document.getElementById('id_emisor_nit').value = data.emisor_nit || '';
                            }
                            
                            const fields = [
                                'tipo_contribuyente', 'regimen_fiscal', 'responsabilidad_tributaria',
                                'pais', 'departamento', 'municipio', 'direccion', 'telefono', 'correo'
                            ];
                            
                            fields.forEach(field => {
                                const el = document.getElementById('id_emisor_' + field);
                                if (el && data['emisor_' + field]) {
                                    el.value = data['emisor_' + field];
                                }
                            });
                        }
                    })
                    .catch(error => {
                        loadingSpan.style.display = 'none';
                        console.error("Error fetching emisor data:", error);
                    });
            }
        });
    }

    if (plantillaInput) {
        plantillaInput.addEventListener('change', function() {
            const val = this.value;
            if (val === 'EXITO') {
                if (emisorNombreInput && (!emisorNombreInput.value || emisorNombreInput.value.includes('CENCOSUD') || emisorNombreInput.value.includes('JUMBO'))) {
                    emisorNombreInput.value = 'ALMACENES ÉXITO SA';
                    if (document.getElementById('id_emisor_nit')) document.getElementById('id_emisor_nit').value = '8909006089';
                    if (serieInput && (!serieInput.value || serieInput.value === 'AU06' || serieInput.value === 'I041')) serieInput.value = 'WJ53';
                }
            } else if (val === 'JUMBO') {
                if (emisorNombreInput && (!emisorNombreInput.value || emisorNombreInput.value.includes('ÉXITO') || emisorNombreInput.value.includes('EXITO'))) {
                    emisorNombreInput.value = 'CENCOSUD COLOMBIA S.A';
                    if (document.getElementById('id_emisor_nit')) document.getElementById('id_emisor_nit').value = '900.155.107-1';
                    if (serieInput && (!serieInput.value || serieInput.value === 'WJ53' || serieInput.value === 'I041')) serieInput.value = 'AU06';
                }
            }
        });
    }
});
