// Esperar a que el DOM esté completamente cargado
document.addEventListener('DOMContentLoaded', () => {
    // Buscar la imagen del logo por su clase o etiqueta
    const logo = document.querySelector('.logo-img');

    if (logo) {
        // Cambiar el cursor a 'puntero' para indicar que es clickeable
        logo.style.cursor = 'pointer';
        
        // Agregar un título hover para mejorar la experiencia de usuario
        logo.title = 'Volver al Inicio';

        // Evento de clic para redirigir al index.html
        logo.addEventListener('click', () => {
            window.location.href = 'index.html';
        });
    }
});