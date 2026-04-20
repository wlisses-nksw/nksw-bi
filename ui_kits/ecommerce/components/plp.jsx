/* NKSW · PLP (Product List Page) — category grid w/ filters */

const PLP = ({ slug, onAdd, onOpenPDP, onGoPLP }) => {
  const [gridCols, setGridCols] = React.useState(4);
  const [openFilter, setOpenFilter] = React.useState(null);
  const [sort, setSort] = React.useState('featured');

  const titles = {
    'new-in': { eyebrow: 'Drop da semana', title: 'New In', sub: 'As novidades da semana — chegaram pra ficar.' },
    'best-sellers': { eyebrow: 'Ícones NKSW', title: 'Best Sellers', sub: 'As peças que não saem de moda.' },
    'biquinis': { eyebrow: 'Categoria', title: 'Biquínis', sub: 'Tops, calcinhas e sets. Mix & match sem regra.' },
    'shop-all': { eyebrow: 'Tudo', title: 'Shop All', sub: 'Toda a curadoria NKSW em um só lugar.' },
    'colecao-stone': { eyebrow: 'Coleção 2026', title: 'Stone', sub: 'Texturas de pedra, tons terrosos. Feita à mão.' },
    'colecao-nacre': { eyebrow: 'Best Sellers', title: 'Nacre', sub: 'Nossa coleção ícone, agora com novos drops.' },
    'maios': { eyebrow: 'Categoria', title: 'Maiôs & Bodies', sub: 'Silhuetas esculturais para o beach club.' },
    'roupas': { eyebrow: 'Categoria', title: 'Roupas', sub: 'Pareôs, vestidos e chemises da curadoria.' },
    'colecoes': { eyebrow: 'Explore', title: 'Coleções', sub: 'Cada coleção NKSW conta uma história.' },
    'sale': { eyebrow: 'Últimos dias', title: 'Summer Sale', sub: 'Até 70% OFF em peças selecionadas.' },
  };
  const meta = titles[slug] || titles['shop-all'];

  const filters = [
    { key: 'categoria', label: 'Categoria', options: ['Tops', 'Calcinhas', 'Maiôs', 'Bodies', 'Vestidos', 'Pareôs'] },
    { key: 'tamanho', label: 'Tamanho', options: ['PP','P','M','G','GG'] },
    { key: 'cor', label: 'Cor', options: ['Milkshake', 'Blueberry', 'Chili', 'Sand Black', 'Nacre', 'Tiger', 'Orbit', 'Nectar'] },
    { key: 'colecao', label: 'Coleção', options: ['Stone', 'Wave', 'Nacre', 'Nectar', 'Muse', 'Origin'] },
    { key: 'preco', label: 'Preço', options: ['até R$ 250', 'R$ 250 – R$ 450', 'R$ 450 – R$ 700', 'R$ 700+'] },
    { key: 'bumbum', label: 'Bumbum', options: ['Tradicional', 'Fio dental'] },
  ];

  const products = slug === 'sale'
    ? PRODUCTS.filter(p => p.compare)
    : PRODUCTS;

  return (
    <main className="nksw-main">
      {/* Breadcrumb */}
      <nav className="nksw-breadcrumb">
        <a onClick={() => onGoPLP('home')}>Home</a>
        <span>/</span>
        <a>Shop</a>
        <span>/</span>
        <strong>{meta.title}</strong>
      </nav>

      {/* Category hero */}
      <section className="nksw-plp-head">
        <div className="t-eyebrow">{meta.eyebrow}</div>
        <h1 className="nksw-plp-title">{meta.title}</h1>
        <p className="nksw-plp-sub">{meta.sub}</p>
      </section>

      {/* Sub-category pills */}
      <div className="nksw-plp-pills">
        {['Todos', 'Tops', 'Calcinhas', 'Sets', 'Maiôs', 'Bodies', 'Pareôs'].map((t,i)=>(
          <button key={i} className={`nksw-pill ${i===0?'is-active':''}`}>{t}</button>
        ))}
      </div>

      {/* Filter bar — sticky */}
      <div className="nksw-plp-bar">
        <div className="nksw-plp-bar-left">
          <button className="nksw-plp-filter-btn"><Icon name="filter" size={16}/> Filtrar & Ordenar</button>
          <span className="t-body-sm" style={{color:'var(--fg-2)'}}>{products.length * 6} peças</span>
        </div>
        <div className="nksw-plp-bar-filters">
          {filters.map(f => (
            <div key={f.key} className="nksw-plp-filter">
              <button onClick={() => setOpenFilter(openFilter === f.key ? null : f.key)}>{f.label} <Icon name="chevronD" size={14}/></button>
              {openFilter === f.key && (
                <div className="nksw-plp-filter-menu">
                  {f.options.map((o,i)=>(
                    <label key={i}><input type="checkbox"/> {o}</label>
                  ))}
                </div>
              )}
            </div>
          ))}
        </div>
        <div className="nksw-plp-bar-right">
          <select className="nksw-plp-sort" value={sort} onChange={e=>setSort(e.target.value)}>
            <option value="featured">Destaque</option>
            <option value="new">Mais recentes</option>
            <option value="asc">Preço · menor</option>
            <option value="desc">Preço · maior</option>
            <option value="bestseller">Mais vendidos</option>
          </select>
          <div className="nksw-plp-view">
            <button onClick={()=>setGridCols(3)} className={gridCols===3?'is-on':''}><Icon name="grid3" size={16}/></button>
            <button onClick={()=>setGridCols(4)} className={gridCols===4?'is-on':''}><Icon name="grid2" size={16}/></button>
          </div>
        </div>
      </div>

      {/* Grid */}
      <section className="nksw-section" style={{paddingTop:24}}>
        <div className={`nksw-pgrid nksw-pgrid-${gridCols}`}>
          {products.map(p => <ProductCard key={p.id} p={p} onAdd={onAdd} onOpenPDP={onOpenPDP} />)}
          {/* Editorial tile inserted in grid */}
          <a className="nksw-grid-editorial">
            <img src="https://images.unsplash.com/photo-1507525428034-b723cf961d3e?w=900&q=80" alt=""/>
            <div className="nksw-grid-editorial-copy">
              <div className="t-eyebrow" style={{color:'#fff'}}>Editorial</div>
              <div style={{fontFamily:'var(--font-serif)', fontSize:28, color:'#fff', fontWeight:500, lineHeight:1.1, marginTop:4}}><em>Ficar bem</em><br/>é conforto.</div>
              <span className="nksw-link-arrow" style={{color:'#fff', marginTop:10}}>GUIA DE FIT <Icon name="arrowR" size={14}/></span>
            </div>
          </a>
          {products.slice(1,4).map(p => <ProductCard key={p.id+'-b'} p={p} onAdd={onAdd} onOpenPDP={onOpenPDP} />)}
        </div>
        <div className="nksw-loadmore">
          <button className="nksw-btn-outline">CARREGAR MAIS · 48 / {products.length*6}</button>
        </div>
      </section>

      <Footer />
    </main>
  );
};

window.PLP = PLP;
