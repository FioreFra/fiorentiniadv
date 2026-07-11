(function () {
  'use strict';

  function escapeHtml(str) {
    const div = document.createElement('div');
    div.textContent = str == null ? '' : str;
    return div.innerHTML;
  }

  function mainPhotoHtml(images, altText) {
    if (!images || !images.length) return 'foto grande outfit';
    return `<img src="${escapeHtml(images[0])}" alt="${escapeHtml(altText)}" loading="lazy" style="position:absolute;inset:0;width:100%;height:100%;object-fit:cover;">`;
  }

  /* Miniatura del capo: se ci sono più immagini (separate da "|" nel foglio),
     mostra una mini-gallery con frecce e puntini invece di una singola foto. */
  function thumbHtml(images, altText, thumbId) {
    if (!images || !images.length) {
      return `<div class="item-thumb placeholder-img">foto capo</div>`;
    }
    if (images.length === 1) {
      return `<div class="item-thumb"><img src="${escapeHtml(images[0])}" alt="${escapeHtml(altText)}" loading="lazy" style="width:100%;height:100%;object-fit:cover;"></div>`;
    }
    const imgs = images.map((url, i) =>
      `<img class="gallery-img${i === 0 ? ' is-active' : ''}" data-index="${i}" src="${escapeHtml(url)}" alt="${escapeHtml(altText)}" loading="lazy">`
    ).join('');
    const dots = images.map((_, i) =>
      `<button type="button" class="gallery-dot${i === 0 ? ' is-active' : ''}" data-dot="${i}" aria-label="Foto ${i + 1}"></button>`
    ).join('');
    return `
      <div class="item-thumb gallery" id="${thumbId}" data-active="0">
        ${imgs}
        <div class="gallery-nav">
          <button type="button" class="gallery-arrow" data-dir="-1" aria-label="Foto precedente">‹</button>
          <button type="button" class="gallery-arrow" data-dir="1" aria-label="Foto successiva">›</button>
        </div>
        <div class="gallery-dots">${dots}</div>
      </div>`;
  }

  function setActiveImage(thumb, index) {
    const imgs = thumb.querySelectorAll('.gallery-img');
    const dots = thumb.querySelectorAll('.gallery-dot');
    const count = imgs.length;
    const next = ((index % count) + count) % count;
    imgs.forEach((img) => img.classList.toggle('is-active', Number(img.dataset.index) === next));
    dots.forEach((dot) => dot.classList.toggle('is-active', Number(dot.dataset.dot) === next));
    thumb.dataset.active = String(next);
  }

  function wireGalleries(container) {
    container.querySelectorAll('.item-thumb.gallery').forEach((thumb) => {
      thumb.querySelectorAll('.gallery-arrow').forEach((btn) => {
        btn.addEventListener('click', () => {
          const current = Number(thumb.dataset.active || '0');
          setActiveImage(thumb, current + Number(btn.dataset.dir));
        });
      });
      thumb.querySelectorAll('.gallery-dot').forEach((dot) => {
        dot.addEventListener('click', () => setActiveImage(thumb, Number(dot.dataset.dot)));
      });
    });
  }

  function renderNotFound() {
    document.querySelector('.look-main').innerHTML =
      '<p style="grid-column:1/-1;padding:24px 0;font:400 15px \'DM Sans\',sans-serif;color:rgba(0,0,0,.6);">Nessun look disponibile al momento. <a href="index.html">Torna alla home</a>.</p>';
  }

  function renderLook(look, products) {
    document.title = `${look.nome} — Sere&You`;
    document.getElementById('crumb-current').textContent = look.nome;
    document.getElementById('look-tag').textContent = look.occasione;
    document.getElementById('look-title').textContent = look.nome;
    document.getElementById('look-caption').textContent = look.caption;

   // link TikTok rimosso: non presente in questa versione del sito

    document.getElementById('look-photo').innerHTML =
      mainPhotoHtml(look.immagini, look.nome) +
      '<span class="look-card__badge">Carosello TikTok</span>';

    const itemList = document.getElementById('item-list');
    const items = look.prodotti_id
      .map((pid) => products[pid])
      .filter(Boolean);

    if (!items.length) {
      itemList.innerHTML = '<p style="padding:16px 0;font:400 14px \'DM Sans\',sans-serif;color:rgba(0,0,0,.6);">Nessun capo disponibile per questo look al momento.</p>';
      return;
    }

    itemList.innerHTML = items.map((item, i) => {
      const thumbId = `item-thumb-${i}`;
      return `
        <div class="item-row">
          ${thumbHtml(item.immagini, item.nome, thumbId)}
          <div>
            <div class="item-name">${escapeHtml(item.nome)}</div>
            <div class="item-price">${escapeHtml(window.SereYouData.formatPrice(item.prezzo))}</div>
            ${item.badge ? `<span class="tag item-badge">${escapeHtml(item.badge)}</span>` : ''}
          </div>
          <a href="${escapeHtml(item.link || '#')}" target="_blank" rel="noopener" class="link-underline item-link">Vedi il prodotto →</a>
        </div>`;
    }).join('');

    wireGalleries(itemList);
  }

  window.addEventListener('DOMContentLoaded', () => {
    window.SereYouData.loadData().then(({ looks, products }) => {
      const ids = Object.keys(looks);
      if (!ids.length) {
        renderNotFound();
        return;
      }
      const requestedId = new URLSearchParams(location.search).get('id');
      const look = (requestedId && looks[requestedId]) ? looks[requestedId] : looks[ids[0]];
      renderLook(look, products);
    });
  });
})();
