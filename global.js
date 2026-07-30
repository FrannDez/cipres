// --- 1. PROTECCIÓN DE RUTAS (CONTROL DE SESIÓN) ---
(function verificarSesion() {
    const usuario = sessionStorage.getItem('usuarioLogueado');
    const esPaginaLogin = window.location.pathname.endsWith('login.html');

    // Si NO está logueado y NO está en la página de login, lo mandamos a loguearse
    if (!usuario && !esPaginaLogin) {
        window.location.href = 'login.html';
    }
})();

// --- 2. COMPORTAMIENTO DEL LOGO GLOBAL ---
document.addEventListener('DOMContentLoaded', () => {
    const logo = document.querySelector('.logo-img');

    if (logo) {
        logo.style.cursor = 'pointer';
        logo.title = 'Volver al Inicio';

        logo.addEventListener('click', () => {
            window.location.href = 'index.html';
        });
    }
});