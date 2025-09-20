/* ---------- Theme handling ---------- */
function initTheme() {
  const theme = localStorage.getItem('theme') || 'dark'; // default = dark
  if (theme === 'light') {
    document.body.classList.add('light-mode');
  } else {
    document.body.classList.remove('light-mode');
  }
}

function toggleTheme() {
  // add a smooth transition class before toggling
  document.body.classList.add('theme-transition');
  const isLight = document.body.classList.toggle('light-mode');
  localStorage.setItem('theme', isLight ? 'light' : 'dark');

  // remove the transition class after animation ends
  setTimeout(() => {
    document.body.classList.remove('theme-transition');
  }, 400); // 400ms matches CSS duration
}

/* ---------- Navigation ---------- */
function goBack() {
  if (window.history.length > 1) {
    window.history.back();
  } else {
    location.href = 'index.html';
  }
}

/* ---------- Counter bar ---------- */
function updateCounterBar(expired, expiring, total) {
  const bar = document.getElementById('appCounters');
  if (!bar) return;
  bar.textContent = `Expired : ${expired} – Expiring : ${expiring} – Total apps : ${total}`;
}

/* ---------- Table helpers ---------- */
function sortTable(idx) {
  const table = document.querySelector('table');
  if (!table) return;

  const tbody = table.tBodies[0];
  const rows = [...tbody.querySelectorAll('tr')];
  const asc = table.dataset.sortColumn == idx && table.dataset.sortOrder !== 'desc';

  table.dataset.sortColumn = idx;
  table.dataset.sortOrder = asc ? 'desc' : 'asc';

  // update header arrows
  table.querySelectorAll('th').forEach((th, i) => {
    th.textContent = th.textContent.replace(/[ ⬆️⬇️]+$/, '');
    if (i === idx) th.textContent += asc ? ' ⬆️' : ' ⬇️';
  });

  rows.sort((a, b) => {
    const A = a.cells[idx].innerText.trim().toLowerCase();
    const B = b.cells[idx].innerText.trim().toLowerCase();
    if (!isNaN(A) && !isNaN(B)) return asc ? A - B : B - A;
    return asc ? A.localeCompare(B) : B.localeCompare(A);
  });

  rows.forEach(r => tbody.appendChild(r));
  recountVisible();
}

function filterTable() {
  const search = document.getElementById('tableSearch')?.value.toLowerCase() || '';
  const status = document.getElementById('statusFilter')?.value || 'all';

  document.querySelectorAll('#appsBody tr').forEach(row => {
    const textMatch = row.innerText.toLowerCase().includes(search);
    const du = parseInt(row.dataset.du, 10);
    let statusMatch = true;

    if (status === 'expired') statusMatch = du <= 0;
    else if (status === 'expiring') statusMatch = du > 0 && du <= 30;

    row.style.display = textMatch && statusMatch ? '' : 'none';
  });

  recountVisible();
}

function recountVisible() {
  const rows = [...document.querySelectorAll('#appsBody tr')].filter(r => r.style.display !== 'none');
  let expired = 0, expiring = 0, total = rows.length;

  rows.forEach(r => {
    const du = parseInt(r.dataset.du, 10);
    if (du <= 0) expired++;
    else if (du <= 30) expiring++;
  });

  updateCounterBar(expired, expiring, total);
}

/* ---------- Footer loader ---------- */
(async () => {
  try {
    const resp = await fetch('footer.html');
    const html = await resp.text();
    const placeholder = document.getElementById('footer-placeholder');
    if (placeholder) placeholder.innerHTML = html;
  } catch (err) {
    console.error('Cannot load footer:', err);
  }
})();

/* ---------- Auto-init ---------- */
window.addEventListener('pageshow', initTheme);
