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

    /* Live Sydney clock — renders into #sydneyClock if present. */
    const _sydTimeFmt = new Intl.DateTimeFormat('en-AU', {
      timeZone: 'Australia/Sydney',
      hour: '2-digit', minute: '2-digit', hour12: false
    });
    const _sydZoneFmt = new Intl.DateTimeFormat('en-AU', {
      timeZone: 'Australia/Sydney', timeZoneName: 'short'
    });
    function tick() {
      const d = new Date();
      const t = _sydTimeFmt.format(d);
      const zoneParts = _sydZoneFmt.formatToParts(d);
      const zone = (zoneParts.find(p => p.type === 'timeZoneName') || {}).value || 'AEST';
      const tEl = document.getElementById('sydneyClock');
      const zEl = document.getElementById('sydneyClockZone');
      if (tEl) tEl.textContent = t;
      if (zEl) zEl.textContent = zone;
    }
    tick();
    setInterval(tick, 30000);
  });
})();
