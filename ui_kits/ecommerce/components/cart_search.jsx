/* NKSW · Cart Drawer, Search Overlay, Sauce widget, Fit banner */

const CartDrawer = ({ open, onClose, items, onRemove, onQty }) => {
  const subtotal = items.reduce((s, i) => s + i.p.price * i.qty, 0);
  const freeShipThreshold = 1200;
  const progress = Math.min(1, subtotal / freeShipThreshold);
  const remaining = Math.max(0, freeShipThreshold - subtotal);

  return (
    <>
      <div className={`nksw-scrim ${open ? 'is-open' : ''}`} onClick={onClose}></div>
      <aside className={`nksw-cart ${open ? 'is-open' : ''}`}>
        <header className="nksw-cart-head">
          <div style={{fontFamily:'var(--font-serif)', fontSize:22, fontWeight:600}}>Sua sacola ({items.length})</div>
          <button className="nksw-icon-btn" onClick={onClose}><Icon name="close" /></button>
        </header>
        <div className="nksw-cart-ship">
          {remaining > 0 ? (
            <div className="t-body-sm">Faltam <strong>{BRL(remaining)}</strong> para <strong>ENVIO GRÁTIS</strong> 📦</div>
          ) : (
            <div className="t-body-sm" style={{color:'var(--ok)'}}>✨ Você ganhou ENVIO GRÁTIS!</div>
          )}
          <div className="nksw-cart-progress"><div style={{width: `${progress*100}%`}}></div></div>
        </div>
        <div className="nksw-cart-items">
          {items.length === 0 && (
            <div style={{textAlign:'center', padding:'48px 20px', color:'var(--fg-2)'}}>
              <Icon name="bag" size={36} />
              <p style={{marginTop:12}}>Sua sacola está vazia.</p>
              <button className="nksw-btn-outline" style={{marginTop:14}} onClick={onClose}>CONTINUAR COMPRANDO</button>
            </div>
          )}
          {items.map(({ p, qty, colorIdx, size }, idx) => (
            <div key={idx} className="nksw-cart-item">
              <img src={p.img} alt=""/>
              <div className="nksw-cart-item-body">
                <div className="nksw-pcard-name" style={{fontSize:12}}>{p.shortName}</div>
                <div className="t-body-sm">{p.colors[colorIdx].name} · Tam {size || 'M'}</div>
                <div style={{display:'flex', justifyContent:'space-between', alignItems:'center', marginTop:10}}>
                  <div className="nksw-qty">
                    <button onClick={() => onQty(idx, qty - 1)}><Icon name="minus" size={14}/></button>
                    <span>{qty}</span>
                    <button onClick={() => onQty(idx, qty + 1)}><Icon name="plus" size={14}/></button>
                  </div>
                  <div className="t-product-price">{BRL(p.price * qty)}</div>
                </div>
                <button className="nksw-cart-remove" onClick={() => onRemove(idx)}>Remover</button>
              </div>
            </div>
          ))}
        </div>
        {items.length > 0 && (
          <footer className="nksw-cart-foot">
            <div className="nksw-cart-totals">
              <div><span>Subtotal</span><span>{BRL(subtotal)}</span></div>
              <div className="t-body-sm"><span>ou 6× de</span><span>{BRL(subtotal/6)} sem juros</span></div>
              <div className="t-body-sm" style={{color:'var(--nksw-red)'}}><span>5% OFF no Pix</span><span>{BRL(subtotal*0.95)}</span></div>
            </div>
            <button className="nksw-btn-primary nksw-btn-block">FINALIZAR COMPRA · {BRL(subtotal)}</button>
            <div className="t-micro" style={{textAlign:'center', marginTop:10, color:'var(--fg-2)'}}>🔒 Pagamento 100% seguro · Frete calculado no checkout</div>
          </footer>
        )}
      </aside>
    </>
  );
};

const SearchOverlay = ({ open, onClose }) => {
  const [q, setQ] = React.useState('');
  const suggested = ['TOP ARIEL', 'Guia de tamanhos', 'Maiô Luna', 'Coleção Stone', 'Calcinha Gaia'];
  const trending = PRODUCTS.slice(0, 4);
  if (!open) return null;
  return (
    <div className="nksw-search-overlay" onClick={(e) => {if (e.target.classList.contains('nksw-search-overlay')) onClose();}}>
      <div className="nksw-search-panel">
        <div className="nksw-search-bar">
          <Icon name="search" size={22} />
          <input autoFocus placeholder="Busque por produto, coleção, cor…" value={q} onChange={e=>setQ(e.target.value)} />
          <button className="nksw-icon-btn" onClick={onClose}><Icon name="close" /></button>
        </div>
        <div className="nksw-search-body">
          <div>
            <div className="t-eyebrow" style={{marginBottom:12}}>Buscas populares</div>
            <div className="nksw-search-chips">
              {suggested.map((s,i)=> <button key={i} className="nksw-chip" onClick={()=>setQ(s)}>{s}</button>)}
            </div>
          </div>
          <div style={{marginTop:28}}>
            <div className="t-eyebrow" style={{marginBottom:16}}>Em alta</div>
            <div className="nksw-search-products">
              {trending.map(p => (
                <a key={p.id} className="nksw-search-prod" href="#">
                  <img src={p.img} alt=""/>
                  <div>
                    <div className="nksw-pcard-name" style={{fontSize:11}}>{p.shortName}</div>
                    <div className="t-body-sm">{BRL(p.price)}</div>
                  </div>
                </a>
              ))}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

/* Sauce-like floating concierge button (bottom right) */
const SauceWidget = () => {
  const [open, setOpen] = React.useState(false);
  return (
    <>
      <button className="nksw-sauce-btn" onClick={() => setOpen(!open)}>
        {open ? <Icon name="close" size={22} /> : <>
          <span style={{fontFamily:'var(--font-serif)', fontStyle:'italic', fontWeight:600, fontSize:20}}>G</span>
        </>}
      </button>
      {open && (
        <div className="nksw-sauce-panel">
          <div className="nksw-sauce-head">
            <div>
              <div style={{fontFamily:'var(--font-serif)', fontSize:18, fontWeight:600}}>Giulia</div>
              <div className="t-body-sm" style={{color:'var(--ok)'}}>● online · concierge NKSW</div>
            </div>
            <button className="nksw-icon-btn" onClick={()=>setOpen(false)}><Icon name="close" size={18} /></button>
          </div>
          <div className="nksw-sauce-body">
            <div className="nksw-bubble">Oi! Sou a Giulia, concierge da NKSW. Como posso te ajudar hoje? ✨</div>
            <div className="nksw-bubble">Posso te ajudar a encontrar o fit perfeito, tirar dúvidas de tamanho, troca ou envio. 🖤</div>
            <div className="nksw-sauce-quick">
              <button>Preciso de ajuda com tamanho</button>
              <button>Status do meu pedido</button>
              <button>Quero uma curadoria para viagem ✈️</button>
            </div>
          </div>
          <div className="nksw-sauce-input">
            <input placeholder="Escreva uma mensagem…" />
            <button><Icon name="arrowR" size={18}/></button>
          </div>
        </div>
      )}
    </>
  );
};

Object.assign(window, { CartDrawer, SearchOverlay, SauceWidget });
