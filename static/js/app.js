// ── Theme toggle ────────────────────────────────────────────────────────────
(function () {
  const btn = document.getElementById('themeToggle');
  if (!btn) return;
  btn.addEventListener('click', function () {
    const isLight = document.documentElement.getAttribute('data-theme') === 'light';
    if (isLight) {
      document.documentElement.removeAttribute('data-theme');
      localStorage.setItem('iu-theme', 'dark');
    } else {
      document.documentElement.setAttribute('data-theme', 'light');
      localStorage.setItem('iu-theme', 'light');
    }
  });
})();

// ── Dashboard: filter chips ──────────────────────────────────────────────────
function filterTable(verdict, btnEl) {
  const rows = document.querySelectorAll('tr.row');

  if (btnEl) {
    document.querySelectorAll('.chip').forEach(c => c.classList.remove('on'));
    btnEl.classList.add('on');
  }

  rows.forEach(row => {
    row.style.display = (verdict === 'all' || row.dataset.verdict === verdict) ? '' : 'none';
  });

  document.querySelectorAll('tr.group-head').forEach(hdr => {
    hdr.style.display = (verdict === 'all' || hdr.dataset.group === verdict) ? '' : 'none';
  });
}

// ── Dashboard: search ────────────────────────────────────────────────────────
function searchTable(q) {
  const query = (q || '').toLowerCase().trim();
  document.querySelectorAll('tr.row').forEach(row => {
    if (!query) { row.style.display = ''; return; }
    const ticker = (row.dataset.ticker || '').toLowerCase();
    const name   = (row.dataset.name   || '').toLowerCase();
    row.style.display = (ticker.includes(query) || name.includes(query)) ? '' : 'none';
  });
}

// ── Dashboard: sort ──────────────────────────────────────────────────────────
let _sortCol = null;
let _sortDir = 1;

function sortTable(col) {
  const tbody = document.getElementById('tableBody');
  if (!tbody) return;
  if (_sortCol === col) { _sortDir *= -1; } else { _sortCol = col; _sortDir = 1; }

  const rows = Array.from(tbody.querySelectorAll('tr.row'));
  rows.sort((a, b) => {
    let av = a.dataset[col] || '';
    let bv = b.dataset[col] || '';
    const an = parseFloat(av);
    const bn = parseFloat(bv);
    if (!isNaN(an) && !isNaN(bn)) return (an - bn) * _sortDir;
    return av.localeCompare(bv) * _sortDir;
  });

  rows.forEach(r => tbody.appendChild(r));
}

// ── Dashboard: row expand ────────────────────────────────────────────────────
document.addEventListener('click', function (e) {
  const row = e.target.closest('tr.row');
  if (!row) return;
  row.classList.toggle('expanded');
});

// ── Qualitative BC/LTP toggle (dashboard) ────────────────────────────────────
function setQual(ticker, field, value, btnEl) {
  // Toggle off if already selected
  const group = btnEl.closest('.qual-group');
  if (group) {
    group.querySelectorAll('.qual-btn').forEach(b => b.classList.remove('active'));
  }
  btnEl.classList.add('active');

  const row = btnEl.closest('tr.row') || btnEl.closest('[data-ticker]');
  const autoScore = row ? parseFloat(row.dataset.score || '0') : 0;

  fetch(`/api/qualitative/${encodeURIComponent(ticker)}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ [field]: value, auto_score: autoScore }),
  })
    .then(r => r.json())
    .then(data => {
      if (data.adj_score != null) {
        const scoreEl = document.querySelector(`[data-ticker="${ticker}"] .adj-score`);
        if (scoreEl) scoreEl.textContent = data.adj_score;
      }
    })
    .catch(err => console.warn('Qual update failed', err));
}

// ── Dashboard: open report ───────────────────────────────────────────────────
function openReport(ticker) {
  window.location.href = `/reports/${ticker}_report.html`;
}
