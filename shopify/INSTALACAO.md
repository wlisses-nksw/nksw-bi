# NKSW · Guia de instalação no Shopify (tema Impulse)

## Estrutura dos arquivos

```
shopify/
├── assets/
│   ├── nksw-tokens.css       ← tokens de design (cores, fontes, espaçamentos)
│   ├── nksw-styles.css       ← estilos de todos os componentes
│   └── nksw-scripts.js       ← JS vanilla (carrinhos, bolinhas, busca…)
├── sections/
│   ├── nksw-announcement-bar.liquid  ← barra de anúncios rotativa
│   ├── nksw-hero.liquid              ← hero editorial
│   ├── nksw-trust-bar.liquid         ← barra de confiança (envio, troca…)
│   ├── nksw-double-collection.liquid ← dupla de coleções
│   ├── nksw-product-grid.liquid      ← grade de produtos (New In, Best Sellers)
│   ├── nksw-editorial-split.liquid   ← editorial "Feita à mão"
│   ├── nksw-triple-tiles.liquid      ← 3 tiles de categoria
│   ├── nksw-instashop.liquid         ← shop the feed Instagram
│   ├── nksw-press-logos.liquid       ← "As Seen On"
│   ├── nksw-newsletter.liquid        ← newsletter 10% OFF
│   └── nksw-buy-together.liquid      ← "Compre junto" (PDP)
├── snippets/
│   ├── nksw-product-card.liquid ← card de produto com bolinhas
│   ├── nksw-cart-drawer.liquid  ← cart drawer lateral
│   └── nksw-sauce-widget.liquid ← concierge Giulia (Sauce)
└── templates/
    └── index.json               ← template da home montado
```

---

## Passo a passo

### 1. Upload dos arquivos

No Shopify Admin → **Temas** → **Editar código** (no tema Impulse):

1. Arraste todos os arquivos de `assets/` para a pasta **Assets**
2. Arraste todos os arquivos de `sections/` para a pasta **Sections**
3. Arraste todos os arquivos de `snippets/` para a pasta **Snippets**
4. Substitua (ou mescle) `templates/index.json`

---

### 2. Adicionar CSS e JS ao tema

Abra `layout/theme.liquid` e adicione **antes de `</head>`**:

```liquid
{{ 'nksw-tokens.css' | asset_url | stylesheet_tag }}
{{ 'nksw-styles.css' | asset_url | stylesheet_tag }}
```

Adicione **antes de `</body>`**:

```liquid
{%- render 'nksw-cart-drawer' -%}
{%- render 'nksw-sauce-widget' -%}
<script src="{{ 'nksw-scripts.js' | asset_url }}" defer></script>
```

---

### 3. Botões que abrem o carrinho e a busca

Nos botões de ícone do header do Impulse, adicione os atributos:

```liquid
{%- comment %} Ícone do carrinho {%- endcomment %}
<button data-open-cart data-cart-count-wrap>
  <!-- ícone do carrinho -->
  <span data-cart-count hidden>0</span>
</button>

{%- comment %} Ícone de busca {%- endcomment %}
<button data-open-search>
  <!-- ícone de busca -->
</button>
```

---

### 4. Bolinhas — metafields dos produtos

As bolinhas trocam o produto exibido no card (não apenas a variante).
Para funcionar, configure os metafields em cada produto:

| Metafield | Namespace | Tipo | Descrição |
|---|---|---|---|
| `color_name` | `custom` | Texto | Nome da cor: "MILKSHAKE" |
| `color_hex` | `custom` | Texto | Hex da cor: "#e4cfb3" |
| `color_siblings` | `custom` | Lista de referências a produto | Produtos-irmãos (mesma peça, outras cores) |
| `companion_1` | `custom` | Referência a produto | 1º produto do "Compre junto" |
| `companion_2` | `custom` | Referência a produto | 2º produto do "Compre junto" |
| `collection_name` | `custom` | Texto | Nome da coleção: "STONE" |

Crie as definições em **Admin → Configurações → Metafields personalizados → Produtos**.

---

### 5. Section "Compre junto" no PDP

Abra `templates/product.json` e adicione a section no `order`:

```json
"nksw-buy-together": {
  "type": "nksw-buy-together",
  "settings": {}
}
```

Os produtos companheiros são definidos via metafields `custom.companion_1` e `custom.companion_2`
em cada produto. Alternativamente, configure globalmente nas settings da section no editor.

---

### 6. Mega menu

O Impulse tem seu próprio mega menu. Para usar o estilo NKSW:

1. No editor do tema, configure as coleções e imagens dentro do menu existente
2. Os estilos `.nksw-mega-*` em `nksw-styles.css` podem ser aplicados ao `nksw-mega-menu`
   custom element — adapte conforme o HTML gerado pelo Impulse.

---

## Notas

- **Fontes** carregadas via Google Fonts (`@import` no `nksw-tokens.css`). Para performance máxima, faça self-host dos arquivos `.woff2` e substitua o `@import`.
- **Sauce**: o snippet `nksw-sauce-widget` é um placeholder visual. Quando o app Sauce estiver instalado, ele injeta seu próprio widget — desabilite o snippet nesse caso.
- **Newsletter**: usa o form nativo do Shopify (`form 'customer'`). O e-mail vai para **Admin → Clientes** com a tag `newsletter`.
- **Desconto bundle** ("Compre junto"): o JS calcula o desconto visualmente. Para aplicá-lo no checkout, configure uma **Função de desconto Shopify** ou use um app de bundle (Bold Bundles, Bundler, etc.).
