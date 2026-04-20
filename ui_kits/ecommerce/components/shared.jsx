/* NKSW · shared primitives used across Home, PLP, PDP, Cart */

const BRL = (n) => 'R$ ' + n.toFixed(2).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.');

/* === Announcement bar + header === */
const AnnouncementBar = () => {
  const msgs = [
    'GARANTA 5% OFF NO PIX',
    'ENVIO GRÁTIS acima de R$ 1.200,00',
    '10% OFF NA SUA PRIMEIRA COMPRA · CÓDIGO: NKSW10',
  ];
  const [i, setI] = React.useState(0);
  React.useEffect(() => { const t = setInterval(() => setI(v => (v + 1) % msgs.length), 4000); return () => clearInterval(t); }, []);
  return (
    <div className="nksw-announce">
      <div className="nksw-announce-track" style={{transform:`translateY(-${i*100}%)`}}>
        {msgs.map((m, k) => <div key={k} className="nksw-announce-msg">{m}</div>)}
      </div>
    </div>
  );
};

const LOGO_URL = '../../assets/logo-nksw.webp';

const Header = ({ onOpenMenu, onOpenSearch, onOpenCart, cartCount, wishCount, scrolled }) => (
  <header className={`nksw-header ${scrolled ? 'is-scrolled' : ''}`}>
    <div className="nksw-header-inner">
      <nav className="nksw-nav-left">
        <NavItem label="SHOP ALL" mega="shopall" onOpenMenu={onOpenMenu} />
        <NavItem label="BIQUÍNIS" mega="biquinis" onOpenMenu={onOpenMenu} />
        <NavItem label="ROUPAS" mega="roupas" onOpenMenu={onOpenMenu} />
        <NavItem label="COLEÇÕES" mega="colecoes" onOpenMenu={onOpenMenu} />
        <NavItem label={<span style={{color:'var(--nksw-red)'}}>SUMMER SALE</span>} mega="sale" onOpenMenu={onOpenMenu} />
      </nav>
      <a href="#" className="nksw-logo" aria-label="NKSW home">
        <img src={LOGO_URL} alt="NKSW" onError={(e)=>{e.target.outerHTML='<span style="font-family:var(--font-serif);font-weight:700;font-size:28px;letter-spacing:.04em">NKSW</span>';}} />
      </a>
      <div className="nksw-nav-right">
        <button className="nksw-icon-btn" onClick={onOpenSearch} aria-label="Buscar"><Icon name="search" /></button>
        <button className="nksw-icon-btn" aria-label="Conta"><Icon name="user" /></button>
        <button className="nksw-icon-btn" aria-label="Wishlist">
          <Icon name="heart" />
          {wishCount > 0 && <span className="nksw-badge">{wishCount}</span>}
        </button>
        <button className="nksw-icon-btn" onClick={onOpenCart} aria-label="Carrinho">
          <Icon name="bag" />
          {cartCount > 0 && <span className="nksw-badge">{cartCount}</span>}
        </button>
      </div>
    </div>
  </header>
);

const NavItem = ({ label, mega, onOpenMenu }) => (
  <div className="nksw-nav-item" onMouseEnter={() => onOpenMenu(mega)}>{label}</div>
);

/* === Icon set (inline SVG — stroke style matches Lucide 1.5) === */
const Icon = ({ name, size = 20 }) => {
  const s = { width: size, height: size, stroke: 'currentColor', fill: 'none', strokeWidth: 1.5, strokeLinecap: 'round', strokeLinejoin: 'round' };
  const paths = {
    search: <><circle cx="11" cy="11" r="7" /><path d="m21 21-4.3-4.3" /></>,
    user: <><circle cx="12" cy="8" r="4" /><path d="M4 21c0-4.4 3.6-8 8-8s8 3.6 8 8" /></>,
    heart: <path d="M20.8 5.6a5.5 5.5 0 0 0-7.8 0L12 6.6l-1-1a5.5 5.5 0 0 0-7.8 7.8l1 1L12 22l7.8-7.6 1-1a5.5 5.5 0 0 0 0-7.8z" />,
    bag: <><path d="M6 7h12l-1 13H7L6 7z" /><path d="M9 7a3 3 0 0 1 6 0" /></>,
    close: <><path d="M18 6 6 18" /><path d="m6 6 12 12" /></>,
    chevronR: <path d="m9 18 6-6-6-6" />,
    chevronL: <path d="m15 18-6-6 6-6" />,
    chevronD: <path d="m6 9 6 6 6-6" />,
    plus: <><path d="M12 5v14" /><path d="M5 12h14" /></>,
    minus: <path d="M5 12h14" />,
    check: <path d="M20 6 9 17l-5-5" />,
    play: <path d="M6 4v16l14-8z" />,
    star: <path d="m12 2 2.9 6.5 7.1.7-5.4 4.8 1.6 7L12 17.3 5.8 21l1.6-7L2 9.2l7.1-.7L12 2z" />,
    truck: <><path d="M1 3h15v13H1z" /><path d="M16 8h4l3 3v5h-7" /><circle cx="6" cy="18" r="2" /><circle cx="18" cy="18" r="2" /></>,
    shield: <path d="M12 2 4 5v7c0 5 3.5 8.5 8 10 4.5-1.5 8-5 8-10V5l-8-3z" />,
    return: <><path d="M3 12a9 9 0 1 0 3-6.7" /><path d="M3 3v6h6" /></>,
    instagram: <><rect x="3" y="3" width="18" height="18" rx="5" /><circle cx="12" cy="12" r="4" /><circle cx="17.5" cy="6.5" r="0.7" fill="currentColor" /></>,
    filter: <><path d="M3 6h18" /><path d="M6 12h12" /><path d="M9 18h6" /></>,
    grid2: <><rect x="3" y="3" width="8" height="8" /><rect x="13" y="3" width="8" height="8" /><rect x="3" y="13" width="8" height="8" /><rect x="13" y="13" width="8" height="8" /></>,
    grid3: <><rect x="3" y="3" width="5" height="5" /><rect x="10" y="3" width="5" height="5" /><rect x="17" y="3" width="4" height="5" /><rect x="3" y="10" width="5" height="5" /><rect x="10" y="10" width="5" height="5" /><rect x="17" y="10" width="4" height="5" /><rect x="3" y="17" width="5" height="4" /><rect x="10" y="17" width="5" height="4" /><rect x="17" y="17" width="4" height="4" /></>,
    whatsapp: <path d="M3 21l1.7-5A8 8 0 1 1 8 19.3L3 21zM8 8c0 0 1 4 4 6 1.5-.3 2-1 2-1" />,
    camera: <><rect x="2" y="6" width="20" height="14" rx="2" /><circle cx="12" cy="13" r="4" /><path d="M8 6V4h8v2" /></>,
    sparkles: <><path d="M12 3v4M12 17v4M3 12h4M17 12h4M6 6l2 2M16 16l2 2M6 18l2-2M16 8l2-2" /></>,
    arrowR: <><path d="M5 12h14" /><path d="m13 6 6 6-6 6" /></>,
  };
  return <svg {...s} viewBox="0 0 24 24">{paths[name]}</svg>;
};

/* === Mega menu === */
const MEGA_DATA = {
  shopall: { cols: [
    { title: 'Featured', links: ['New in', 'Best Sellers', 'Clássicos', 'Naked Atelier'] },
    { title: 'Shop by', links: ['Biquínis', 'Maiôs & Bodies', 'Roupas', 'Acessórios'] },
    { title: 'Coleções', links: ['Stone', 'Wave', 'Nacre', 'Nectar', 'Muse'] },
  ], image: 'https://images.unsplash.com/photo-1493558103817-58b2924bce98?w=1200&q=80', imgLabel: 'NEW · STONE' },
  biquinis: { cols: [
    { title: 'Tops', links: ['Todos os Tops', 'Ariel', 'Annie', 'Yara', 'Isla', 'Amara', 'Nina'] },
    { title: 'Calcinhas', links: ['Todas as Calcinhas', 'Gaia', 'Alana', 'Loren', 'Perla', 'Dahlia'] },
    { title: 'Maiôs & Bodies', links: ['Todos', 'Body Lisa', 'Body Drapeado'] },
  ], image: 'https://images.unsplash.com/photo-1582719508461-905c673771fd?w=1200&q=80', imgLabel: 'TOP ANNIE — BRISA' },
  roupas: { cols: [
    { title: 'Tops & Camisas', links: ['Shirts', 'Tops', 'Chemises'] },
    { title: 'Bottoms', links: ['Calças', 'Saias'] },
    { title: 'Praia', links: ['Pareôs', 'Vestidos', 'Acessórios'] },
  ], image: 'https://images.unsplash.com/photo-1539109136881-3be0616acf4b?w=1200&q=80', imgLabel: 'VESTIDO SOLE' },
  colecoes: { cols: [
    { title: 'Drops 2026', links: ['Stone', 'Wave', 'Nacre'] },
    { title: 'Ícones', links: ['Origin', 'Muse', 'Nectar', 'Milkshake', 'Blueberry'] },
    { title: 'Statement', links: ['Tiger', 'Leopard', 'Orbit', 'Sand Black'] },
  ], image: 'https://images.unsplash.com/photo-1507525428034-b723cf961d3e?w=1200&q=80', imgLabel: 'COLLECTION · NACRE' },
  sale: { cols: [
    { title: 'Summer Sale até 70% OFF', links: ['70% OFF', '60% OFF', '50% OFF', '40% OFF', '30% OFF'] },
    { title: 'Promo destacada', links: ['Best Sellers Promo', 'Final de coleção', 'Últimas peças'] },
  ], image: 'https://nksw.co/cdn/shop/files/Foto_19-04-2025_20_16_35.jpg?v=1769770563&width=1200', imgLabel: 'BEST SELLERS — 40% OFF' },
};

const MegaMenu = ({ open, onClose }) => {
  if (!open) return null;
  const data = MEGA_DATA[open];
  if (!data) return null;
  return (
    <div className="nksw-mega" onMouseLeave={onClose}>
      <div className="nksw-mega-inner">
        <div className="nksw-mega-cols">
          {data.cols.map((c, i) => (
            <div key={i} className="nksw-mega-col">
              <div className="t-eyebrow" style={{marginBottom: '12px'}}>{c.title}</div>
              <ul>{c.links.map((l, j) => <li key={j}><a href="#">{l}</a></li>)}</ul>
            </div>
          ))}
        </div>
        <a href="#" className="nksw-mega-feature" style={{backgroundImage: `url(${data.image})`}}>
          <div className="nksw-mega-feature-overlay">
            <div className="t-eyebrow" style={{color:'#fff',opacity:.85}}>Em destaque</div>
            <div style={{fontFamily:'var(--font-serif)', fontSize:28, fontWeight:600, color:'#fff', lineHeight:1.1, marginTop:6}}>{data.imgLabel}</div>
            <span className="nksw-link-arrow" style={{color:'#fff', marginTop:14}}>VER TUDO <Icon name="arrowR" size={14}/></span>
          </div>
        </a>
      </div>
    </div>
  );
};

Object.assign(window, { BRL, AnnouncementBar, Header, Icon, MegaMenu });
