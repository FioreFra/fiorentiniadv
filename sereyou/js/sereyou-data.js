(function (global) {
  'use strict';

  /* ═══════════════════════════════════════════════════════════
     CONFIGURAZIONE — incolla qui i due link CSV pubblicati da
     Google Sheets:
       Google Sheets → File → Condividi → Pubblica sul web
       → scegli il singolo foglio ("Outfit" o "Prodotti")
       → formato CSV → Pubblica → copia il link.
     Finché uno dei due resta col valore placeholder, il sito
     continua a funzionare con i dati di esempio qui sotto.
     ═══════════════════════════════════════════════════════════ */
  const CONFIG = {
    outfitsCsvUrl: 'INCOLLA_QUI_IL_LINK_CSV_OUTFITS',
    productsCsvUrl: 'INCOLLA_QUI_IL_LINK_CSV_PRODOTTI',
  };

  const SAMPLE_PRODUCTS = {
    p01: { nome: 'Blazer strutturato color panna', prezzo: 79.90, immagini: [], link: '#', badge: '' },
    p02: { nome: 'T-shirt basic girocollo', prezzo: 14.90, immagini: [], link: '#', badge: '' },
    p03: { nome: 'Pantalone sartoriale a palazzo', prezzo: 54.90, immagini: [], link: '#', badge: '' },
    p04: { nome: 'Maglione oversize mix cachemire', prezzo: 62, immagini: [], link: '#', badge: 'Ultimo pezzo' },
    p05: { nome: 'Sneakers bianche minimal', prezzo: 39, immagini: [], link: '#', badge: '' },
    p06: { nome: 'Jeans dritto vita alta', prezzo: 48, immagini: [], link: '#', badge: '' },
  };

  const SAMPLE_LOOKS = {
    l1: {
      nome: 'Look Ufficio Cool', occasione: 'Ufficio smart', immagini: [],
      caption: 'Questo è il mio "ufficio ma non voglio sembrare in ufficio". Il blazer fa tutto il lavoro, sotto ci sta comoda anche una t-shirt semplice.',
      link_carosello: '#', prodotti_id: ['p01', 'p02', 'p03'],
    },
    l2: {
      nome: 'Look Aperitivo Easy', occasione: 'Aperitivo', immagini: [],
      caption: "Il look che metto quando esco dal lavoro e vado dritta all'aperitivo, senza passare da casa.",
      link_carosello: '#', prodotti_id: ['p04', 'p05'],
    },
    l3: {
      nome: 'Look Weekend Comfy', occasione: 'Weekend fuori porta', immagini: [],
      caption: 'Il look che metto quando devo camminare tutto il giorno e non voglio pensarci.',
      link_carosello: '#', prodotti_id: ['p05', 'p06'],
    },
  };

  function isConfigured() {
    return /^https?:\/\//.test(CONFIG.outfitsCsvUrl) && /^https?:\/\//.test(CONFIG.productsCsvUrl);
  }

  function parseCsv(url) {
    return new Promise((resolve, reject) => {
      Papa.parse(url, {
        download: true,
        header: true,
        skipEmptyLines: true,
        complete: (res) => resolve(res.data),
        error: reject,
      });
    });
  }

  function isActive(row) {
    return (row.attivo || '').trim().toUpperCase() === 'SI';
  }

  function imageList(cell) {
    return (cell || '').split('|').map((s) => s.trim()).filter(Boolean);
  }

  function toPrice(raw) {
    return parseFloat((raw || '0').toString().replace(',', '.')) || 0;
  }

  function formatPrice(n) {
    return '€ ' + n.toFixed(2).replace('.', ',');
  }

  async function loadData() {
    if (!isConfigured()) {
      return { products: SAMPLE_PRODUCTS, looks: SAMPLE_LOOKS, usingFallback: true };
    }
    try {
      const [productRows, outfitRows] = await Promise.all([
        parseCsv(CONFIG.productsCsvUrl),
        parseCsv(CONFIG.outfitsCsvUrl),
      ]);

      const products = {};
      productRows.forEach((r) => {
        const id = (r.id || '').trim();
        if (!id || !isActive(r)) return;
        products[id] = {
          nome: (r.nome || '').trim(),
          prezzo: toPrice(r.prezzo),
          immagini: imageList(r.immagine_url),
          link: (r.link_affiliato || '#').trim(),
          badge: (r.badge || '').trim(),
        };
      });

      const looks = {};
      outfitRows.forEach((r) => {
        const id = (r.id || '').trim();
        if (!id || !isActive(r)) return;
        looks[id] = {
          nome: (r.nome || '').trim(),
          occasione: (r.occasione || '').trim(),
          immagini: imageList(r.immagine_url),
          caption: (r.caption || '').trim(),
          link_carosello: (r.link_carosello || '#').trim(),
          prodotti_id: (r.prodotti_id || '').split(',').map((s) => s.trim()).filter(Boolean),
        };
      });

      return { products, looks, usingFallback: false };
    } catch (err) {
      console.error('Sere&You: errore nel caricamento dei CSV, uso i dati di esempio.', err);
      return { products: SAMPLE_PRODUCTS, looks: SAMPLE_LOOKS, usingFallback: true };
    }
  }

  global.SereYouData = { loadData, imageList, formatPrice };
})(window);
