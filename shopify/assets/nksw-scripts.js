/* ============================================================
   NKSW Scripts — Shopify Impulse
   Vanilla JS, sem dependências externas
   ============================================================ */

const BRL = (cents) =>
  (cents / 100).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });

/* ── Announcement Bar ───────────────────────────────────────── */
class NKSWAnnouncementBar extends HTMLElement {
  connectedCallback() {
    this.track = this.querySelector('.nksw-announce-track');
    this.items = this.querySelectorAll('.nksw-announce-msg');
    this.current = 0;
    if (this.items.length > 1) {
      setInterval(() => {
        this.current = (this.current + 1) % this.items.length;
        this.track.style.transform = `translateY(-${this.current * 100}%)`;
      }, 4000);
    }
  }
}
customElements.define('nksw-announcement-bar', NKSWAnnouncementBar);

/* ── Mega Menu ──────────────────────────────────────────────── */
class NKSWMegaMenu extends HTMLElement {
  connectedCallback() {
    const triggers = document.querySelectorAll('[data-mega]');
    const panels   = this.querySelectorAll('[data-mega-panel]');
    let closeTimer;

    const show = (key) => {
      clearTimeout(closeTimer);
      panels.forEach(p => p.hidden = p.dataset.megaPanel !== key);
      this.hidden = false;
    };
    const hide = () => { closeTimer = setTimeout(() => { this.hidden = true; }, 180); };

    triggers.forEach(t => {
      t.addEventListener('mouseenter', () => show(t.dataset.mega));
      t.addEventListener('focus',      () => show(t.dataset.mega));
    });
    this.addEventListener('mouseenter', () => clearTimeout(closeTimer));
    this.addEventListener('mouseleave', hide);
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape') { this.hidden = true; } });
  }
}
customElements.define('nksw-mega-menu', NKSWMegaMenu);

/* ── Product Card — bolinhas / swatch hover ─────────────────── */
class NKSWProductCard extends HTMLElement {
  connectedCallback() {
    this.mainImg  = this.querySelector('.nksw-pcard-img--main');
    this.hoverImg = this.querySelector('.nksw-pcard-img--hover');
    this.colorEl  = this.querySelector('.nksw-pcard-color');
    this.priceEl  = this.querySelector('.nksw-pcard-prices');
    this.linkEls  = this.querySelectorAll('a[data-pcard-link]');
    this.wishBtn  = this.querySelector('.nksw-pcard-wish');

    // Hover crossfade
    this.addEventListener('mouseenter', () => this.classList.add('is-hovered'));
    this.addEventListener('mouseleave', () => this.classList.remove('is-hovered'));

    // Wishlist toggle
    this.wishBtn?.addEventListener('click', (e) => {
      e.preventDefault();
      this.wishBtn.classList.toggle('is-on');
    });

    // Swatch dots
    this.querySelectorAll('.nksw-dot[data-handle]').forEach(dot => {
      dot.addEventListener('mouseenter', () => this._loadSibling(dot));
    });
  }

  async _loadSibling(dot) {
    const handle = dot.dataset.handle;
    if (!handle) return;

    // Mesmo produto — volta ao estado base
    if (handle === this.dataset.handle) {
      this._resetToBase();
      this._setActive(dot);
      return;
    }

    // Cache hit
    if (dot._data) { this._apply(dot, dot._data); return; }

    try {
      const res  = await fetch(`/products/${handle}.js`);
      if (!res.ok) return;
      const data = await res.json();
      dot._data  = data;
      this._apply(dot, data);
    } catch (_) {}
  }

  _apply(dot, data) {
    this._setActive(dot);
    const img1 = data.images[0], img2 = data.images[1] || img1;
    if (img1 && this.mainImg)  this.mainImg.src  = img1 + '&width=800';
    if (img2 && this.hoverImg) this.hoverImg.src = img2 + '&width=800';
    if (this.colorEl) this.colorEl.textContent = ' · ' + (dot.dataset.colorName || '');

    if (this.priceEl) {
      const p = data.price, c = data.compare_at_price;
      if (c && c > p) {
        const disc = Math.round(100 - (p / c) * 100);
        this.priceEl.innerHTML =
          `<span class="nksw-price-was">${BRL(c)}</span>` +
          `<span class="nksw-price-sale">${BRL(p)}</span>` +
          `<span class="nksw-pcard-badge nksw-pcard-badge--discount" style="position:static;margin-left:6px">-${disc}%</span>`;
      } else {
        this.priceEl.innerHTML = `<span class="nksw-price">${BRL(p)}</span>`;
      }
    }

    const url = `/products/${data.handle}`;
    this.linkEls.forEach(l => l.href = url);
    const quickBtn = this.querySelector('[data-quickadd-trigger]');
    if (quickBtn) quickBtn.dataset.url = url;
    this.querySelectorAll('[data-quickadd-variant]').forEach((btn, i) => {
      if (data.variants[i]) btn.dataset.quickaddVariant = data.variants[i].id;
    });
  }

  _setActive(dot) {
    this.querySelectorAll('.nksw-dot').forEach(d => d.classList.remove('is-sel'));
    dot.classList.add('is-sel');
  }

  _resetToBase() {
    const d = this.dataset;
    if (this.mainImg  && d.img)   this.mainImg.src  = d.img;
    if (this.hoverImg && d.img2)  this.hoverImg.src = d.img2;
    if (this.colorEl  && d.color) this.colorEl.textContent = ' · ' + d.color;
    if (this.priceEl  && d.priceHtml) this.priceEl.innerHTML = d.priceHtml;
    this.linkEls.forEach(l => l.href = d.url || '#');
  }
}
customElements.define('nksw-product-card', NKSWProductCard);

/* ── Quick Add ──────────────────────────────────────────────── */
class NKSWQuickAdd extends HTMLElement {
  connectedCallback() {
    const trigger = this.querySelector('[data-quickadd-trigger]');
    const panel   = this.querySelector('.nksw-quickadd-panel');

    trigger?.addEventListener('click', (e) => {
      e.stopPropagation();
      e.preventDefault();
      panel?.classList.toggle('is-open');
    });

    this.querySelectorAll('[data-quickadd-variant]').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        panel?.classList.remove('is-open');
        document.dispatchEvent(new CustomEvent('nksw:cart:add', {
          detail: { variantId: +btn.dataset.quickaddVariant }
        }));
      });
    });

    document.addEventListener('click', () => panel?.classList.remove('is-open'));
  }
}
customElements.define('nksw-quick-add', NKSWQuickAdd);

/* ── Cart Drawer ────────────────────────────────────────────── */
class NKSWCartDrawer extends HTMLElement {
  connectedCallback() {
    this.inner  = this.querySelector('.nksw-cart-inner');
    this.scrim  = document.getElementById('nksw-scrim');
    this.counts = document.querySelectorAll('[data-cart-count]');

    document.addEventListener('nksw:cart:open', () => this.open());
    document.addEventListener('nksw:cart:add',  (e) => this.addItem(e.detail));

    this.querySelector('[data-cart-close]')?.addEventListener('click', () => this.close());
    this.scrim?.addEventListener('click', () => this.close());

    this.addEventListener('click', (e) => {
      const removeBtn = e.target.closest('[data-cart-remove]');
      if (removeBtn) { e.preventDefault(); this.changeLine(+removeBtn.dataset.cartRemove, 0); }

      const qtyBtn = e.target.closest('[data-cart-qty]');
      if (qtyBtn) {
        const item  = qtyBtn.closest('[data-line]');
        const line  = +item.dataset.line;
        const delta = qtyBtn.dataset.cartQty === '+' ? 1 : -1;
        const cur   = +item.dataset.qty;
        this.changeLine(line, cur + delta);
      }
    });

    // Carrega contagem inicial sem abrir
    fetch('/cart.js').then(r => r.json()).then(c => this._updateCount(c.item_count));
  }

  open() {
    this.setAttribute('open', '');
    this.scrim?.classList.add('is-open');
    document.body.style.overflow = 'hidden';
    this.render();
  }

  close() {
    this.removeAttribute('open');
    this.scrim?.classList.remove('is-open');
    document.body.style.overflow = '';
  }

  async addItem({ variantId, quantity = 1, properties = {} }) {
    const btn = document.querySelector('[data-atc-btn]');
    if (btn) { btn.disabled = true; btn.textContent = 'ADICIONANDO...'; }
    try {
      const res = await fetch('/cart/add.js', {
        method:  'POST',
        headers: { 'Content-Type': 'application/json' },
        body:    JSON.stringify({ id: variantId, quantity, properties })
      });
      if (!res.ok) {
        const err = await res.json();
        alert(err.description || 'Produto esgotado neste tamanho.');
        return;
      }
      await this.render();
      this.open();
    } catch (e) {
      alert('Erro ao adicionar ao carrinho. Tente novamente.');
    } finally {
      if (btn) { btn.disabled = false; btn.textContent = 'ADICIONAR AO CARRINHO'; }
    }
  }

  async changeLine(line, quantity) {
    await fetch('/cart/change.js', {
      method:  'POST',
      headers: { 'Content-Type': 'application/json' },
      body:    JSON.stringify({ line, quantity })
    });
    await this.render();
  }

  async render() {
    const cart = await fetch('/cart.js').then(r => r.json());
    this._updateCount(cart.item_count);

    const THRESHOLD = 120000; // R$ 1.200 em centavos
    const remaining = Math.max(0, THRESHOLD - cart.total_price);
    const progress  = Math.min(1, cart.total_price / THRESHOLD) * 100;

    const shipMsg = remaining > 0
      ? `Faltam <strong>${BRL(remaining)}</strong> para <strong>ENVIO GRÁTIS</strong> 📦`
      : `<span style="color:var(--ok)">✨ Você ganhou ENVIO GRÁTIS!</span>`;

    let itemsHTML = '';
    if (cart.items.length === 0) {
      itemsHTML = `
        <div class="nksw-cart-empty">
          <svg width="40" height="40" fill="none" stroke="currentColor" stroke-width="1.5" viewBox="0 0 24 24">
            <path d="M6 7h12l-1 13H7L6 7z"/><path d="M9 7a3 3 0 0 1 6 0"/>
          </svg>
          <p>Sua sacola está vazia.</p>
          <button class="nksw-btn-outline" style="margin-top:8px" data-cart-close>CONTINUAR COMPRANDO</button>
        </div>`;
    } else {
      cart.items.forEach((item, i) => {
        const line    = i + 1;
        const variant = item.variant_title !== 'Default Title' ? item.variant_title : '';
        itemsHTML += `
          <div class="nksw-cart-item" data-line="${line}" data-qty="${item.quantity}">
            <a href="${item.url}">
              <img src="${item.image}" alt="${item.product_title}" loading="lazy"/>
            </a>
            <div class="nksw-cart-item-body">
              <div class="nksw-cart-item-title">${item.product_title}</div>
              ${variant ? `<div class="nksw-cart-item-variant">${variant}</div>` : ''}
              <div class="nksw-cart-item-bottom">
                <div class="nksw-qty">
                  <button data-cart-qty="-" aria-label="Diminuir">−</button>
                  <span>${item.quantity}</span>
                  <button data-cart-qty="+" aria-label="Aumentar">+</button>
                </div>
                <span class="nksw-cart-item-price">${BRL(item.line_price)}</span>
              </div>
              <button class="nksw-cart-remove" data-cart-remove="${line}">Remover</button>
            </div>
          </div>`;
      });
    }

    let footerHTML = '';
    if (cart.items.length > 0) {
      footerHTML = `
        <footer class="nksw-cart-foot">
          <div class="nksw-cart-totals">
            <div class="nksw-cart-row"><span>Subtotal</span><span>${BRL(cart.total_price)}</span></div>
            <div class="nksw-cart-row nksw-cart-row--sm"><span>ou 6× de</span><span>${BRL(Math.round(cart.total_price / 6))} sem juros</span></div>
            <div class="nksw-cart-row nksw-cart-row--sm nksw-cart-row--pix"><span>5% OFF no Pix</span><span>${BRL(Math.round(cart.total_price * 0.95))}</span></div>
          </div>
          <a href="/checkout" class="nksw-btn nksw-btn-primary nksw-btn-block">
            FINALIZAR COMPRA · ${BRL(cart.total_price)}
          </a>
          <p class="nksw-cart-secure">🔒 Pagamento 100% seguro · Frete calculado no checkout</p>
        </footer>`;
    }

    this.inner.innerHTML = `
      <div class="nksw-cart-ship">
        <div class="nksw-cart-ship-msg">${shipMsg}</div>
        <div class="nksw-cart-progress"><div style="width:${progress}%"></div></div>
      </div>
      <div class="nksw-cart-items">${itemsHTML}</div>
      ${footerHTML}`;
  }

  _updateCount(n) {
    this.counts.forEach(el => {
      el.textContent = n;
      el.hidden = n === 0;
    });
  }
}
customElements.define('nksw-cart-drawer', NKSWCartDrawer);

/* ── Search Overlay ─────────────────────────────────────────── */
class NKSWSearchOverlay extends HTMLElement {
  connectedCallback() {
    this.input    = this.querySelector('input[type=search]');
    this.results  = this.querySelector('[data-search-results]');
    this._timer   = null;

    document.addEventListener('nksw:search:open', () => this.open());
    this.querySelector('[data-search-close]')?.addEventListener('click', () => this.close());
    this.addEventListener('click', (e) => { if (e.target === this) this.close(); });
    document.addEventListener('keydown', (e) => { if (e.key === 'Escape') this.close(); });

    this.input?.addEventListener('input', () => {
      clearTimeout(this._timer);
      this._timer = setTimeout(() => this._search(this.input.value.trim()), 280);
    });
  }

  open() {
    this.setAttribute('open', '');
    setTimeout(() => this.input?.focus(), 60);
  }

  close() {
    this.removeAttribute('open');
    if (this.input)   this.input.value = '';
    if (this.results) this.results.innerHTML = '';
  }

  async _search(q) {
    if (!this.results) return;
    if (q.length < 2)  { this.results.innerHTML = ''; return; }
    try {
      const url  = `/search/suggest.json?q=${encodeURIComponent(q)}&resources[type]=product&resources[limit]=6`;
      const data = await fetch(url).then(r => r.json());
      const products = data.resources?.results?.products || [];
      this.results.innerHTML = products.map(p => `
        <a href="${p.url}" class="nksw-search-prod">
          <img src="${p.image}" alt="${p.title}" loading="lazy"/>
          <div>
            <div class="nksw-search-prod-name">${p.title}</div>
            <div class="nksw-search-prod-price">${BRL(p.price)}</div>
          </div>
        </a>`).join('');
    } catch (_) {}
  }
}
customElements.define('nksw-search-overlay', NKSWSearchOverlay);

/* ── Sauce / Giulia ─────────────────────────────────────────── */
class NKSWSauceWidget extends HTMLElement {
  connectedCallback() {
    const btn   = this.querySelector('.nksw-sauce-btn');
    const panel = this.querySelector('.nksw-sauce-panel');

    btn?.addEventListener('click', () => {
      const open = panel?.classList.toggle('is-open');
      btn.setAttribute('aria-expanded', String(!!open));
    });
    this.querySelector('[data-sauce-close]')?.addEventListener('click', () => {
      panel?.classList.remove('is-open');
      btn?.setAttribute('aria-expanded', 'false');
    });
    this.querySelectorAll('[data-sauce-action]').forEach(b => {
      b.addEventListener('click', () => {
        const inp = this.querySelector('.nksw-sauce-input input');
        if (inp) { inp.value = b.textContent.trim(); inp.focus(); }
      });
    });
  }
}
customElements.define('nksw-sauce-widget', NKSWSauceWidget);

/* ── Buy Together ───────────────────────────────────────────── */
class NKSWBuyTogether extends HTMLElement {
  connectedCallback() {
    this._update();
    this.querySelectorAll('input[type=checkbox]').forEach(cb =>
      cb.addEventListener('change', () => this._update())
    );
    this.querySelector('[data-bt-add]')?.addEventListener('click', () => this._addAll());
  }

  _update() {
    const cards = this.querySelectorAll('.nksw-bt-card');
    let raw = 0, count = 0;

    cards.forEach(card => {
      const cb  = card.querySelector('input[type=checkbox]');
      const sel = cb?.checked;
      card.classList.toggle('is-off', !sel);
      if (sel) { raw += +card.dataset.price; count++; }
    });

    const disc  = count === 3 ? 0.10 : count === 2 ? 0.05 : 0;
    const total = raw * (1 - disc);
    const saved = raw - total;

    const set = (sel, val) => {
      const el = this.querySelector(sel);
      if (el) el.textContent = val;
    };
    const setHTML = (sel, val) => {
      const el = this.querySelector(sel);
      if (el) el.innerHTML = val;
    };

    // Preços em reais (já estão em reais no data-price, não em centavos)
    const fmt = (v) => v.toLocaleString('pt-BR', { style:'currency', currency:'BRL' });

    set('[data-bt-subtotal]', fmt(raw));
    set('[data-bt-total]',    fmt(total));
    set('[data-bt-installment]', fmt(total / 6));
    set('[data-bt-count]',   `${count}/3 peças`);

    const discRow = this.querySelector('[data-bt-discount-row]');
    if (discRow) {
      discRow.hidden = disc === 0;
      set('[data-bt-discount-label]', count === 3 ? 'Desconto bundle 10%' : 'Desconto duo 5%');
      set('[data-bt-saved]', `− ${fmt(saved)}`);
    }

    setHTML('[data-bt-hint]', count === 3
      ? `<strong>Você está economizando ${fmt(saved)}</strong> levando o conjunto completo.`
      : count === 2
        ? `Adicione a terceira peça e leve <strong>10% OFF</strong> em todo o conjunto.`
        : `Monte seu conjunto: duas peças garantem <strong>5% OFF</strong>, três garantem <strong>10% OFF</strong>.`);

    const btn = this.querySelector('[data-bt-add]');
    if (btn) {
      btn.disabled    = count === 0;
      btn.textContent = count === 0
        ? 'SELECIONE AO MENOS UMA PEÇA'
        : `ADICIONAR ${count === 1 ? 'PEÇA' : 'CONJUNTO'} · ${fmt(total)}`;
    }
  }

  _addAll() {
    this.querySelectorAll('.nksw-bt-card').forEach(card => {
      const cb = card.querySelector('input[type=checkbox]');
      if (!cb?.checked) return;
      const variantId = card.dataset.variantId;
      if (variantId) {
        document.dispatchEvent(new CustomEvent('nksw:cart:add', {
          detail: { variantId: +variantId }
        }));
      }
    });
  }
}
customElements.define('nksw-buy-together', NKSWBuyTogether);

/* ── Triggers globais ───────────────────────────────────────── */
document.addEventListener('click', (e) => {
  if (e.target.closest('[data-open-cart]'))   document.dispatchEvent(new CustomEvent('nksw:cart:open'));
  if (e.target.closest('[data-open-search]')) document.dispatchEvent(new CustomEvent('nksw:search:open'));
});
