// static/js/operador-dashboard.js
document.addEventListener('DOMContentLoaded', function() {
    function checkNuevosOperador() {
        fetch('/operador/api/nuevos')
            .then(response => response.json())
            .then(data => {
                // Actualizar el contador de nuevos (si existe el elemento)
                const contadorNuevos = document.querySelector('.contador-nuevos');
                if (contadorNuevos) {
                    if (data.nuevos > 0) {
                        contadorNuevos.innerHTML = `<i class="fas fa-bell"></i> Tienes <strong>${data.nuevos}</strong> pedidos nuevos por revisar`;
                        contadorNuevos.style.display = 'flex';
                    } else {
                        contadorNuevos.style.display = 'none';
                    }
                }

                // Actualizar el badge de "Nuevo" en las tarjetas de pedidos
                const tarjetas = document.querySelectorAll('.glass-card .badge-nuevo');
                // Si hay pedidos nuevos, mostrar/ocultar según corresponda
                // Pero esto requiere que el backend tenga lógica para saber qué pedidos son nuevos
                // Alternativa: recargar la lista de pedidos (más complejo)
                // Por simplicidad, solo actualizamos el contador de nuevos.
            })
            .catch(error => console.error('Error al verificar nuevos pedidos:', error));
    }

    // Ejecutar cada 10 segundos
    setInterval(checkNuevosOperador, 10000);
    // Ejecutar al cargar
    checkNuevosOperador();
});