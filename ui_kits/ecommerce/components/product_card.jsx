/* NKSW · ProductCard — grid-layout vitrine card
   · image 3:4 sem radius
   · hover: crossfade para img2 + overlay "ADICIONAR AO CARRINHO"
   · bolinhas de cor: hover troca a imagem do card (padrão Monday Swimwear)
*/
const ProductCard = ({ p, onAdd, onOpenPDP, dense = false }) => {
  const [hover, setHover] = React.useState(false);
  const [colorIdx, setColorIdx] = React.useState(p.colors.findIndex(c => c.sel) || 0);
  const [wished, setWished] = React.useState(false);
  const hasSale = !!p.compare;
  const discount = hasSale ? Math.round(100 - (p.price / p.compare) * 100) : 0;

  return (
    <div className={`nksw-pcard ${dense ? 'is-dense' : ''}`} onMouseEnter={() => setHover(true)} onMouseLeave={() => setHover(false)}>
      <div className="nksw-pcard-media" onClick={() => onOpenPDP(p)}>
        <img src={p.img} alt={p.name} className={`nksw-pcard-img ${hover ? 'is-hover' : ''}`} loading="lazy" />
        <img src={p.img2} alt="" className={`nksw-pcard-img nksw-pcard-img-2 ${hover ? 'is-hover' : ''}`} loading="lazy" />
        {p.badge && <span className={`nksw-badge-tag nksw-badge-${p.badge.toLowerCase().replace(/\s+/g,'-')}`}>{p.badge}</span>}
        {hasSale && <span className="nksw-badge-tag nksw-badge-discount">-{discount}%</span>}
        <button className={`nksw-wish-btn ${wished ? 'is-on' : ''}`} onClick={(e) => {e.stopPropagation(); setWished(!wished);}} aria-label="Wishlist">
          <Icon name="heart" size={18} />
        </button>
        <div className={`nksw-pcard-quickadd ${hover ? 'is-visible' : ''}`}>
          <button onClick={(e) => {e.stopPropagation(); onAdd(p, colorIdx);}}>ADICIONAR AO CARRINHO</button>
        </div>
      </div>
      <div className="nksw-pcard-info">
        <div className="nksw-pcard-name">{p.shortName}<span className="nksw-pcard-color"> · {p.colors[colorIdx].name}</span></div>
        <div className="nksw-pcard-prices">
          {hasSale ? (<>
            <span className="t-product-price-was">{BRL(p.compare)}</span>
            <span className="t-product-price-sale">{BRL(p.price)}</span>
          </>) : (
            <span className="t-product-price">{BRL(p.price)}</span>
          )}
        </div>
        <div className="nksw-pcard-dots">
          {p.colors.map((c, i) => (
            <button
              key={i}
              onMouseEnter={() => setColorIdx(i)}
              className={`nksw-dot ${i === colorIdx ? 'is-sel' : ''}`}
              style={{background: c.hex, borderColor: c.hex === '#ffffff' || c.hex.toLowerCase() === '#fff' ? '#e5e7eb' : 'transparent'}}
              aria-label={c.name}
            />
          ))}
          {p.colors.length > 4 && <span className="nksw-dot-more">+{p.colors.length - 4}</span>}
        </div>
      </div>
    </div>
  );
};

window.ProductCard = ProductCard;
