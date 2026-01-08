(() => {
  const container = document.getElementById('sidebar-container');
  if (!container) return;

  const currentPath = window.location.pathname.split('/').pop() || 'index.html';

  fetch('sidebar.html')
    .then((res) => {
      if (!res.ok) {
        throw new Error('Sidebar load failed');
      }
      return res.text();
    })
    .then((html) => {
      container.innerHTML = html;
      const links = Array.from(container.querySelectorAll('[data-nav]'));
      links.forEach((link) => {
        if (link.getAttribute('data-nav') === currentPath) {
          link.classList.add('bg-blue-600', 'text-white', 'font-semibold');
        }
      });
    })
    .catch(() => {
      container.innerHTML = '';
    });
})();
