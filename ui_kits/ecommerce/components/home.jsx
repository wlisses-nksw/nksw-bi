/* NKSW · Home */
const Home = ({ onAdd, onOpenPDP, onGoPLP }) => {
  const [email, setEmail] = React.useState('');

  return (
    <main className="nksw-main">
      {/* HERO */}
      <section className="nksw-hero">
        <img src="https://images.unsplash.com/photo-1507525428034-b723cf961d3e?w=2000&q=80" alt="" />
        <div className="nksw-hero-overlay">
          <div className="t-eyebrow" style={{color:'#fff', letterSpacing:'.22em'}}>VERÃO 2026 · DROP STONE</div>
          <h1 className="nksw-hero-title">Dare to be<br/><em>yourself.</em></h1>
          <div style={{display:'flex', gap:32, marginTop:32, alignItems:'center'}}>
            <button className="nksw-btn-editorial" onClick={() => onGoPLP('shop-all')}>SHOP NEW IN →</button>
            <button className="nksw-btn-editorial nksw-btn-editorial-ghost" onClick={() => onGoPLP('colecoes')}>DESCUBRA AS COLEÇÕES</button>
          </div>
        </div>
      </section>

      {/* TRUST BAR */}
      <section className="nksw-trust">
        <div className="nksw-trust-item"><Icon name="truck" /><span>ENVIO GRÁTIS acima de R$ 1.200</span></div>
        <div className="nksw-trust-item"><Icon name="return" /><span>TROCA FÁCIL em troquefacil.com.br</span></div>
        <div className="nksw-trust-item"><Icon name="shield" /><span>ATÉ 6× SEM JUROS · 5% OFF NO PIX</span></div>
        <div className="nksw-trust-item"><Icon name="sparkles" /><span>FEITO À MÃO EM BRASÍLIA</span></div>
      </section>

      {/* DOUBLE FEATURE — COLLECTIONS */}
      <section className="nksw-double">
        <a className="nksw-double-card" onClick={() => onGoPLP('colecao-stone')}>
          <img src="https://images.unsplash.com/photo-1493558103817-58b2924bce98?w=1400&q=80" alt=""/>
          <div className="nksw-double-overlay">
            <div className="t-eyebrow" style={{color:'#fff'}}>NOVA COLEÇÃO</div>
            <div className="nksw-double-title">STONE</div>
            <span className="nksw-link-arrow" style={{color:'#fff'}}>SHOP STONE <Icon name="arrowR" size={14}/></span>
          </div>
        </a>
        <a className="nksw-double-card" onClick={() => onGoPLP('colecao-nacre')}>
          <img src="https://images.unsplash.com/photo-1515886657613-9f3515b0c78f?w=1400&q=80" alt=""/>
          <div className="nksw-double-overlay">
            <div className="t-eyebrow" style={{color:'#fff'}}>BEST SELLERS</div>
            <div className="nksw-double-title">NACRE</div>
            <span className="nksw-link-arrow" style={{color:'#fff'}}>SHOP NACRE <Icon name="arrowR" size={14}/></span>
          </div>
        </a>
      </section>

      {/* NEW IN — product carousel */}
      <section className="nksw-section">
        <div className="nksw-section-head">
          <div>
            <div className="t-eyebrow">Drop da semana</div>
            <h2 className="nksw-section-title">New In</h2>
          </div>
          <a className="nksw-link-arrow" onClick={() => onGoPLP('new-in')}>VER TUDO <Icon name="arrowR" size={14}/></a>
        </div>
        <div className="nksw-pgrid nksw-pgrid-4">
          {PRODUCTS.slice(0, 4).map(p => <ProductCard key={p.id} p={p} onAdd={onAdd} onOpenPDP={onOpenPDP} />)}
        </div>
      </section>

      {/* EDITORIAL SPLIT — "Made in Brazil" warm sand background */}
      <section className="nksw-editorial">
        <div className="nksw-editorial-media">
          <img src="https://images.unsplash.com/photo-1582719508461-905c673771fd?w=1200&q=80" alt=""/>
        </div>
        <div className="nksw-editorial-copy">
          <div className="t-eyebrow">Desde 2016 · Brasília · Brasil</div>
          <h2 className="nksw-editorial-title"><em>Curadoria</em><br/>feita à mão.</h2>
          <p>Cada peça NKSW é desenvolvida com modelagem exclusiva e tecidos selecionados para entregar o fit perfeito. Da praia de Trancoso ao beach club de Ibiza — nossa curadoria acompanha sua próxima história. ✨</p>
          <a className="nksw-link-arrow" onClick={() => onGoPLP('sobre')}>NOSSA HISTÓRIA <Icon name="arrowR" size={14}/></a>
        </div>
      </section>

      {/* BESTSELLERS grid */}
      <section className="nksw-section">
        <div className="nksw-section-head">
          <div>
            <div className="t-eyebrow">Ícones NKSW</div>
            <h2 className="nksw-section-title">Best sellers</h2>
          </div>
          <a className="nksw-link-arrow" onClick={() => onGoPLP('best-sellers')}>VER TUDO <Icon name="arrowR" size={14}/></a>
        </div>
        <div className="nksw-pgrid nksw-pgrid-4">
          {PRODUCTS.slice(2, 6).map(p => <ProductCard key={p.id} p={p} onAdd={onAdd} onOpenPDP={onOpenPDP} />)}
        </div>
      </section>

      {/* TRIPLE CATEGORY TILES */}
      <section className="nksw-triple">
        {[
          {label:'BIQUÍNIS', sub:'Tops, calcinhas, sets', img:'https://images.unsplash.com/photo-1566174053879-31528523f8ae?w=1000&q=80', slug:'biquinis'},
          {label:'MAIÔS & BODIES', sub:'Silhuetas esculturais', img:'https://images.unsplash.com/photo-1507525428034-b723cf961d3e?w=1000&q=80', slug:'maios'},
          {label:'ROUPAS', sub:'Pareôs, vestidos, chemises', img:'https://images.unsplash.com/photo-1539109136881-3be0616acf4b?w=1000&q=80', slug:'roupas'},
        ].map((c,i) => (
          <a key={i} className="nksw-triple-card" onClick={() => onGoPLP(c.slug)}>
            <img src={c.img} alt=""/>
            <div className="nksw-triple-overlay">
              <div className="nksw-triple-label">{c.label}</div>
              <div className="nksw-triple-sub">{c.sub}</div>
            </div>
          </a>
        ))}
      </section>

      {/* INSTASHOP */}
      <section className="nksw-section">
        <div className="nksw-section-head" style={{textAlign:'center', justifyContent:'center', flexDirection:'column', gap:6}}>
          <div className="t-eyebrow">@nakedswimwear · #NakedBabes</div>
          <h2 className="nksw-section-title" style={{textAlign:'center'}}>Shop the feed</h2>
        </div>
        <div className="nksw-instashop">
          {[1,2,3,4,5,6].map(i => (
            <a key={i} className="nksw-insta-tile" href="#">
              <img src={[PRODUCTS[0].img, PRODUCTS[1].img, PRODUCTS[2].img, PRODUCTS[3].img, PRODUCTS[4].img, PRODUCTS[5].img][i-1]} alt=""/>
              <div className="nksw-insta-overlay"><Icon name="instagram" size={18} /></div>
            </a>
          ))}
        </div>
      </section>

      {/* AS SEEN ON */}
      <section className="nksw-press">
        <div className="t-eyebrow" style={{textAlign:'center'}}>AS SEEN ON</div>
        <div className="nksw-press-logos">
          <span style={{fontFamily:'var(--font-serif)', fontWeight:700, fontSize:28}}>VOGUE</span>
          <span style={{fontFamily:'var(--font-serif)', fontStyle:'italic', fontWeight:500, fontSize:22}}>Marie Claire</span>
          <span style={{fontFamily:'var(--font-sans)', fontWeight:700, letterSpacing:'.16em', fontSize:18}}>GLAMOUR</span>
          <span style={{fontFamily:'var(--font-serif)', fontWeight:600, fontSize:26}}>Forbes</span>
          <span style={{fontFamily:'var(--font-serif)', fontStyle:'italic', fontWeight:600, fontSize:22}}>Harper's Bazaar</span>
        </div>
      </section>

      {/* NEWSLETTER */}
      <section className="nksw-newsletter">
        <div className="t-eyebrow" style={{color:'#fff', opacity:.8}}>Let's get naked!</div>
        <h2 style={{fontFamily:'var(--font-serif)', fontSize:48, fontWeight:500, color:'#fff', margin:'14px 0 8px', lineHeight:1.1}}>10% OFF na sua primeira compra.</h2>
        <p style={{color:'rgba(255,255,255,.75)', fontSize:15, maxWidth:520, margin:'0 auto 28px'}}>Receba lançamentos em primeira mão, acesso antecipado a drops e curadoria semanal direto no seu e-mail.</p>
        <form className="nksw-news-form" onSubmit={e=>{e.preventDefault(); alert('Welcome to NKSW ✨');}}>
          <input type="email" required value={email} onChange={e=>setEmail(e.target.value)} placeholder="seu.email@exemplo.com" />
          <button type="submit">QUERO 10% OFF</button>
        </form>
      </section>

      <Footer />
    </main>
  );
};

const Footer = () => (
  <footer className="nksw-footer">
    <div className="nksw-footer-top">
      <div className="nksw-footer-brand">
        <div style={{fontFamily:'var(--font-serif)', fontWeight:700, fontSize:36, letterSpacing:'.04em'}}>NKSW</div>
        <p className="t-body-sm" style={{maxWidth:260, marginTop:12}}>Moda praia premium feita em Brasília. Para quem não tem tempo para segunda opção.</p>
        <div className="nksw-social">
          <a href="#" aria-label="Instagram"><Icon name="instagram"/></a>
          <a href="#" aria-label="WhatsApp"><Icon name="whatsapp"/></a>
          <a href="#" aria-label="TikTok"><Icon name="camera"/></a>
        </div>
      </div>
      <div className="nksw-footer-col">
        <div className="t-eyebrow">Shop</div>
        <ul><li>Biquínis</li><li>Maiôs & Bodies</li><li>Roupas</li><li>Acessórios</li><li>Summer Sale</li></ul>
      </div>
      <div className="nksw-footer-col">
        <div className="t-eyebrow">Ajuda</div>
        <ul><li>Atendimento</li><li>Trocas & devoluções</li><li>Guia de tamanhos</li><li>Rastreio de pedido</li><li>FAQ</li></ul>
      </div>
      <div className="nksw-footer-col">
        <div className="t-eyebrow">Institucional</div>
        <ul><li>Nossa história</li><li>Lojas físicas</li><li>Sustentabilidade</li><li>Imprensa</li><li>Termos & privacidade</li></ul>
      </div>
    </div>
    <div className="nksw-footer-bottom">
      <div className="t-micro">© 2026 NAKED SWIMWEAR · CNPJ XX.XXX.XXX/0001-XX · BRASÍLIA, BR</div>
      <div className="nksw-pay">
        <PayLogo>VISA</PayLogo><PayLogo>MASTER</PayLogo><PayLogo>AMEX</PayLogo><PayLogo>PIX</PayLogo><PayLogo>SHOP</PayLogo>
      </div>
    </div>
  </footer>
);

const PayLogo = ({ children }) => (
  <span style={{border:'1px solid #e5e7eb', borderRadius:4, padding:'4px 8px', fontSize:10, fontWeight:700, letterSpacing:'.08em', color:'#6b7280', background:'#fff'}}>{children}</span>
);

window.Home = Home;
