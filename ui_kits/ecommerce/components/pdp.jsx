/* NKSW · PDP (Product Detail Page) */

/* === BuyTogether — "Compre junto" ===
   Editorial bundle: current product + 2 suggested pieces, toggleable,
   with a bundle-discount reveal. Calm composition, no iconographic slop. */
const BuyTogether = ({ main, onAdd, onOpenPDP }) => {
  // Pick two complementary pieces: prefer a bottom + a cover-up, skip the main.
  const pool = PRODUCTS.filter(x => x.id !== main.id);
  const companion1 = pool.find(x => /CALCINHA|BODY/i.test(x.name)) || pool[0];
  const companion2 = pool.find(x => /VESTIDO|MAIÔ/i.test(x.name) && x.id !== companion1.id) || pool[1];
  const initial = [main, companion1, companion2];

  const [sel, setSel] = React.useState({ [main.id]: true, [companion1.id]: true, [companion2.id]: true });
  const toggle = (id) => setSel(s => ({...s, [id]: !s[id]}));

  const selected = initial.filter(p => sel[p.id]);
  const rawTotal = selected.reduce((s,p) => s + p.price, 0);
  const allThree = selected.length === 3;
  const discount = allThree ? 0.10 : selected.length === 2 ? 0.05 : 0;
  const total = rawTotal * (1 - discount);
  const saved = rawTotal - total;

  const addBundle = () => {
    selected.forEach(p => onAdd(p, p.colors.findIndex(c => c.sel) >= 0 ? p.colors.findIndex(c => c.sel) : 0, 'M'));
  };

  return (
    <section className="nksw-section nksw-buytogether">
      <div className="nksw-bt-head">
        <div className="t-eyebrow">Styled together</div>
        <h2 className="nksw-section-title">Compre junto</h2>
        <p className="nksw-bt-sub">
          Selecionado pela equipe NKSW para compor o seu <em>{main.collection}</em>.
          {' '}Leve os três e ganhe <strong>10% OFF</strong> no conjunto.
        </p>
      </div>

      <div className="nksw-bt-stage">
        <div className="nksw-bt-row">
          {initial.map((prod, i) => (
            <React.Fragment key={prod.id}>
              <label className={`nksw-bt-card ${sel[prod.id] ? 'is-on' : 'is-off'} ${prod.id === main.id ? 'is-main' : ''}`}>
                <input type="checkbox" checked={!!sel[prod.id]} onChange={() => toggle(prod.id)} aria-label={`Incluir ${prod.shortName}`}/>
                <div className="nksw-bt-card-img" onClick={(e) => { e.preventDefault(); onOpenPDP(prod); }}>
                  <img src={prod.img} alt={prod.name} />
                  {prod.id === main.id && <span className="nksw-bt-pin">Esta peça</span>}
                  <span className="nksw-bt-check" aria-hidden>
                    {sel[prod.id] ? <Icon name="check" size={14}/> : null}
                  </span>
                </div>
                <div className="nksw-bt-card-info">
                  <div className="t-micro" style={{color:'var(--fg-2)', letterSpacing:'.12em'}}>{prod.collection}</div>
                  <div className="nksw-bt-card-name">{prod.shortName}</div>
                  <div className="nksw-bt-card-price">{BRL(prod.price)}</div>
                </div>
              </label>
              {i < initial.length - 1 && (
                <div className="nksw-bt-plus" aria-hidden>+</div>
              )}
            </React.Fragment>
          ))}
        </div>

        <aside className="nksw-bt-summary">
          <div className="nksw-bt-sum-head">
            <span className="t-eyebrow">Seu conjunto</span>
            <span className="nksw-bt-count">{selected.length}/3 peças</span>
          </div>

          <ul className="nksw-bt-sum-list">
            {initial.map(prod => (
              <li key={prod.id} className={sel[prod.id] ? '' : 'is-off'}>
                <span>{prod.shortName}</span>
                <span>{BRL(prod.price)}</span>
              </li>
            ))}
          </ul>

          <div className="nksw-bt-sum-totals">
            <div className="nksw-bt-sum-row">
              <span>Subtotal</span>
              <span>{BRL(rawTotal)}</span>
            </div>
            {discount > 0 && (
              <div className="nksw-bt-sum-row nksw-bt-sum-save">
                <span>Desconto {allThree ? 'bundle 10%' : 'duo 5%'}</span>
                <span>− {BRL(saved)}</span>
              </div>
            )}
            <div className="nksw-bt-sum-row nksw-bt-sum-total">
              <span>Total</span>
              <div style={{textAlign:'right'}}>
                <div style={{fontFamily:'var(--font-serif)', fontSize:28, fontWeight:500, lineHeight:1}}>{BRL(total)}</div>
                <div className="t-micro" style={{color:'var(--fg-2)', marginTop:4}}>ou 6× de {BRL(total/6)} sem juros</div>
              </div>
            </div>
          </div>

          <button
            className="nksw-btn-primary nksw-btn-block"
            style={{marginTop:18}}
            disabled={selected.length === 0}
            onClick={addBundle}
          >
            {selected.length === 0 ? 'SELECIONE AO MENOS UMA PEÇA' : `ADICIONAR ${selected.length === 1 ? 'PEÇA' : 'CONJUNTO'} · ${BRL(total)}`}
          </button>

          <div className="nksw-bt-hint">
            {allThree
              ? <><strong>Você está economizando {BRL(saved)}</strong> levando o conjunto completo.</>
              : selected.length === 2
                ? <>Adicione a terceira peça e leve <strong>10% OFF</strong> em todo o conjunto.</>
                : <>Monte seu conjunto: duas peças garantem <strong>5% OFF</strong>, três garantem <strong>10% OFF</strong>.</>
            }
          </div>
        </aside>
      </div>
    </section>
  );
};

const PDP = ({ p, onAdd, onGoPLP, onOpenPDP }) => {
  const [colorIdx, setColorIdx] = React.useState(p.colors.findIndex(c => c.sel) || 0);
  const [size, setSize] = React.useState(null);
  const [accord, setAccord] = React.useState('desc');
  const [thumbIdx, setThumbIdx] = React.useState(0);
  const images = [p.img, p.img2, p.img, p.img2];

  return (
    <main className="nksw-main">
      <nav className="nksw-breadcrumb">
        <a onClick={() => onGoPLP('home')}>Home</a>
        <span>/</span>
        <a onClick={() => onGoPLP('biquinis')}>Biquínis</a>
        <span>/</span>
        <a onClick={() => onGoPLP('colecao-stone')}>{p.collection}</a>
        <span>/</span>
        <strong>{p.shortName}</strong>
      </nav>

      <section className="nksw-pdp">
        <div className="nksw-pdp-gallery">
          <div className="nksw-pdp-thumbs">
            {images.map((src, i) => (
              <button key={i} className={`nksw-pdp-thumb ${i===thumbIdx?'is-active':''}`} onClick={()=>setThumbIdx(i)}>
                <img src={src} alt=""/>
              </button>
            ))}
          </div>
          <div className="nksw-pdp-main-img">
            <img src={images[thumbIdx]} alt={p.name} />
            {p.badge && <span className={`nksw-badge-tag nksw-badge-${p.badge.toLowerCase().replace(/\s+/g,'-')}`}>{p.badge}</span>}
          </div>
        </div>

        <div className="nksw-pdp-info">
          <div className="t-eyebrow">COLEÇÃO · {p.collection}</div>
          <h1 className="nksw-pdp-title">{p.shortName}</h1>
          <div className="nksw-pdp-rating">
            {[1,2,3,4,5].map(i => <Icon key={i} name="star" size={14} />)}
            <span style={{color:'var(--fg-2)', marginLeft:6, fontSize:13}}>4.9 · 127 avaliações</span>
          </div>

          <div className="nksw-pdp-price">
            {p.compare ? (<>
              <span className="t-product-price-was" style={{fontSize:16}}>{BRL(p.compare)}</span>
              <span className="t-product-price-sale" style={{fontSize:24}}>{BRL(p.price)}</span>
              <span className="nksw-badge-tag nksw-badge-discount" style={{position:'static'}}>-{Math.round(100-(p.price/p.compare)*100)}%</span>
            </>) : (
              <span style={{fontSize:24, fontWeight:500}}>{BRL(p.price)}</span>
            )}
          </div>
          <div className="t-body-sm" style={{color:'var(--fg-2)'}}>
            ou 6× de <strong>{BRL(p.price/6)}</strong> sem juros · <span style={{color:'var(--nksw-red)', fontWeight:600}}>5% OFF no Pix: {BRL(p.price*0.95)}</span>
          </div>

          {/* Size */}
          <div className="nksw-pdp-section">
            <div className="nksw-pdp-label">
              <span>TAMANHO: <strong>{size || '—'}</strong></span>
            </div>
            <div className="nksw-pdp-sizes">
              {p.sizes.map(s => (
                <button key={s} onClick={()=>setSize(s)} className={`nksw-size-sw ${size===s?'is-sel':''}`}>{s}</button>
              ))}
            </div>
          </div>

          {/* Bumbum selector (calcinha pattern) */}
          {p.category === 'calcinha' && (
            <div className="nksw-pdp-section">
              <div className="nksw-pdp-label">BUMBUM</div>
              <div className="nksw-pdp-sizes">
                <button className="nksw-size-sw is-sel" style={{minWidth:120}}>TRADICIONAL</button>
                <button className="nksw-size-sw" style={{minWidth:120}}>FIO DENTAL</button>
              </div>
            </div>
          )}

          {/* CTA */}
          <button className="nksw-btn-primary nksw-btn-block nksw-pdp-add" onClick={() => onAdd(p, colorIdx, size || 'M')}>
            {size ? 'ADICIONAR AO CARRINHO' : 'SELECIONE UM TAMANHO'}
          </button>
          <button className="nksw-btn-outline nksw-btn-block" style={{marginTop:10}}>
            <Icon name="sparkles" size={16}/> PROVADOR VIRTUAL
          </button>

          {/* Trust list */}
          <ul className="nksw-pdp-trust">
            <li><Icon name="truck" size={16}/> <span><strong>Envio em até 48h</strong> · Grátis acima de R$ 1.200</span></li>
            <li><Icon name="return" size={16}/> <span><strong>Troca fácil</strong> · Primeira troca por nossa conta</span></li>
            <li><Icon name="sparkles" size={16}/> <span><strong>Feito em Brasília</strong> · Modelagem exclusiva</span></li>
          </ul>

          {/* Accordions */}
          <div className="nksw-pdp-accord">
            {[
              {key:'desc', title:'Descrição', body: (<p>Top estilo cropped com detalhe de drapeado no busto — modelagem exclusiva NKSW que acompanha o movimento do corpo. Tecido encorpado com proteção UV. Ideal para quem busca silhueta marcada e sustentação sem usar bojo.</p>)},
              {key:'guide', title:'Guia de tamanhos', body: (<>
                <p style={{marginBottom:12}}>Use a tabela abaixo como referência. Em caso de dúvida entre dois tamanhos, recomendamos o <strong>maior</strong> para tops e o <strong>menor</strong> para calcinhas.</p>
                <table className="nksw-size-table">
                  <thead>
                    <tr><th>Tamanho</th><th>Busto</th><th>Cintura</th><th>Quadril</th></tr>
                  </thead>
                  <tbody>
                    <tr><td>PP</td><td>80–84</td><td>60–64</td><td>86–90</td></tr>
                    <tr><td>P</td><td>84–88</td><td>64–68</td><td>90–94</td></tr>
                    <tr><td>M</td><td>88–92</td><td>68–72</td><td>94–98</td></tr>
                    <tr><td>G</td><td>92–96</td><td>72–76</td><td>98–102</td></tr>
                    <tr><td>GG</td><td>96–100</td><td>76–80</td><td>102–106</td></tr>
                  </tbody>
                </table>
                <p className="t-micro" style={{color:'var(--fg-2)', marginTop:10}}>Medidas em centímetros. A modelagem NKSW é <strong>true to size</strong>.</p>
              </>)},
              {key:'fit', title:'Caimento & Fit', body: (<>
                <p><strong>Modelagem:</strong> Slim com regulagem nas alças.</p>
                <p><strong>Medidas da modelo:</strong> 1,72m · usando tam P.</p>
                <p><strong>Indicado para:</strong> Busto pequeno a médio.</p>
              </>)},
              {key:'tec', title:'Composição & Cuidados', body: (<>
                <p><strong>Composição:</strong> 94,2% Poliamida · 5,8% Elastano · dupla camada.</p>
                <p><strong>Cuidados:</strong> Lavar à mão com água fria. Não torcer. Secar à sombra. Não passar.</p>
              </>)},
              {key:'env', title:'Envio & Trocas', body: (<p>Enviamos para todo o Brasil via Correios e transportadoras parceiras. Prazo médio: SP/RJ 2–4 dias úteis, demais capitais 3–5, interior 5–8. Trocas em até 30 dias via <strong>nakedswimwear.troquefacil.com.br</strong>.</p>)},
            ].map(a => (
              <div key={a.key} className={`nksw-acc ${accord===a.key?'is-open':''}`}>
                <button onClick={() => setAccord(accord === a.key ? '' : a.key)}>
                  <span>{a.title}</span>
                  <Icon name={accord===a.key?'minus':'plus'} size={16}/>
                </button>
                {accord===a.key && <div className="nksw-acc-body">{a.body}</div>}
              </div>
            ))}
          </div>
        </div>
      </section>

      {/* Compre junto — frequently bought together · delicate, curated */}
      <BuyTogether main={p} onAdd={onAdd} onOpenPDP={onOpenPDP} />

      {/* Shop the look */}
      <section className="nksw-section">
        <div className="nksw-section-head" style={{justifyContent:'center', textAlign:'center', flexDirection:'column', gap:6}}>
          <div className="t-eyebrow">Complete the look</div>
          <h2 className="nksw-section-title" style={{textAlign:'center'}}>Shop the look</h2>
        </div>
        <div className="nksw-pgrid nksw-pgrid-4">
          {PRODUCTS.filter(x => x.id !== p.id).slice(0,4).map(p2 => <ProductCard key={p2.id} p={p2} onAdd={onAdd} onOpenPDP={onOpenPDP}/>)}
        </div>
      </section>

      {/* Reviews */}
      <section className="nksw-section nksw-reviews">
        <div className="nksw-reviews-head">
          <div>
            <div className="t-eyebrow">O que dizem as NakedBabes</div>
            <h2 className="nksw-section-title">Avaliações (127)</h2>
          </div>
          <div className="nksw-reviews-score">
            <div style={{fontFamily:'var(--font-serif)', fontSize:56, fontWeight:500}}>4.9</div>
            <div>{[1,2,3,4,5].map(i=><Icon key={i} name="star" size={16}/>)}</div>
            <div className="t-body-sm" style={{color:'var(--fg-2)'}}>baseado em 127 avaliações</div>
          </div>
        </div>
        <div className="nksw-reviews-grid">
          {[
            {name:'Mariana L.', t:'Ficou perfeito', body:'Caimento impecável, sustenta muito bem e a cor Milkshake é ainda mais linda pessoalmente. Já é meu 3º biquíni NKSW.'},
            {name:'Beatriz R.', t:'Amei', body:'Levei para Trancoso, não saí dele. O tecido não amassa mesmo depois de horas na praia.'},
            {name:'Luisa F.', t:'Indico!', body:'Uso P no geral e pedi P — serviu certinho. Qualidade premium de verdade.'},
          ].map((r,i)=>(
            <div key={i} className="nksw-review">
              <div>{[1,2,3,4,5].map(s => <Icon key={s} name="star" size={12}/>)}</div>
              <div style={{fontFamily:'var(--font-serif)', fontSize:20, fontWeight:600, margin:'10px 0 6px'}}>{r.t}</div>
              <p className="t-body-sm">{r.body}</p>
              <div className="t-micro" style={{marginTop:12, color:'var(--fg-2)'}}>— {r.name} · cliente verificado</div>
            </div>
          ))}
        </div>
      </section>

      <Footer />
    </main>
  );
};

window.PDP = PDP;
