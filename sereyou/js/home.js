(function () {
  'use strict';

  function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str == null ? '' : str;
    return div.innerHTML;
  }

  function mediaHtml(images, altText) {
    if (!images || !images.length) return 'foto outfit';
    return `<img src="${escapeHtml(images[0])}" alt="${escapeHtml(altText)}" loading="lazy" style="position:absolute;inset:0;width:100%;height:100%;object-fit:cover;">`;
  }

  function renderLooks(looks) {
    const grid = document.getElementById('looks-grid');
    const countEl = document.getElementById('looks-count');
    const ids = Object.keys(looks);

    countEl.textContent = ids.length
      ? `${ids.length} look · aggiornati ogni settimana`
      : 'aggiornati ogni settimana';

    if (!ids.length) {
      grid.innerHTML = '<p style="grid-column:1/-1;padding:32px;font:400 14px \'DM Sans\',sans-serif;color:rgba(0,0,0,.6);">Nessun look pubblicato al momento. Torna a trovarci presto!</p>';
      return;
    }
    grid.innerHTML = ids.map((id) => {
  const look = looks[id];
  return `
    <a href="look.html?id=${encodeURIComponent(id)}" class="look-card">
      <div class="look-card__media placeholder-img">
        ${mediaHtml(look.immagini, look.nome)}
        <span class="look-card__badge">Carosello TikTok</span>
      </div>
      <div class="look-card__body">
        <span class="tag">${escapeHtml(look.occasione)}</span>
        <h3 class="look-card__title">${escapeHtml(look.nome)}</h3>
        <span class="link-underline look-card__link">Guarda il look →</span>
      </div>
    </a>`;
}).join('');

  }

  window.addEventListener('DOMContentLoaded', () => {
    window.SereYouData.loadData().then(({ looks }) => renderLooks(looks));
  });
})();
