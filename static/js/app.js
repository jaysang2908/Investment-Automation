/* Equity.Research — shared site JS */
(function () {
  /* Theme init is done inline in <head> to avoid flash.
     This file handles the toggle interaction and the live clock. */

  document.addEventListener('DOMContentLoaded', function () {
    const root = document.documentElement;

    /* Theme toggle */
    const tog = document.getElementById('themeToggle');
    if (tog) {
      tog.addEventListener('click', function () {
        const next = root.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
        root.setAttribute('data-theme', next);
        try { localStorage.setItem('eq.theme', next); } catch(e) {}
      });
    }

    /* Live UTC clock */
    function tick() {
      const d = new Date();
      const h = String(d.getUTCHours()).padStart(2, '0');
      const m = String(d.getUTCMinutes()).padStart(2, '0');
      const el = document.getElementById('navTime');
      if (el) el.textContent = h + ':' + m + ' UTC';
    }
    tick();
    setInterval(tick, 30000);
  });
})();
