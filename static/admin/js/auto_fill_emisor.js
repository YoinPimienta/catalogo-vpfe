document.addEventListener('DOMContentLoaded', function() {
    const emisorNombreInput = document.getElementById('id_emisor_nombre');
    
    if (emisorNombreInput) {
        // Create an indicator for loading state
        const loadingSpan = document.createElement('span');
        loadingSpan.id = 'emisor-loading';
        loadingSpan.style.display = 'none';
        loadingSpan.style.marginLeft = '10px';
        loadingSpan.style.color = '#666';
        loadingSpan.textContent = 'Buscando datos...';
        emisorNombreInput.parentNode.appendChild(loadingSpan);

        emisorNombreInput.addEventListener('blur', function() {
            const nombre = this.value.trim();
            if (nombre) {
                loadingSpan.style.display = 'inline';
                fetch(`/api/emisor-data/?emisor_nombre=${encodeURIComponent(nombre)}`)
                    .then(response => response.json())
                    .then(data => {
                        loadingSpan.style.display = 'none';
                        if (Object.keys(data).length > 0) {
                            if (!document.getElementById('id_emisor_nit').value) {
                                document.getElementById('id_emisor_nit').value = data.emisor_nit || '';
                            }
                            
                            // Define fields and logic to avoid overwriting existing data if preferred
                            // But usually autofill overwrites since the user is typing the name
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
});
