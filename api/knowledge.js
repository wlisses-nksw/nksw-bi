/**
 * NKSW Knowledge Base API
 * Retorna toda a base de conhecimento estática da Naked Swimwear.
 * O agente da Voll consulta este endpoint para responder sobre
 * políticas, medidas, cupons, contatos e identidade da marca.
 */

const knowledge = {
  marca: {
    nome: "Naked Swimwear (NKSW)",
    slogan: "Ouse ser você mesma",
    descricao:
      "Marca brasileira de moda praia de luxo acessível, fundada em Brasília por Ananda e Camilla. Produção 100% feita à mão no Brasil, com DNA global e foco em empoderar mulheres de todos os corpos.",
    valores: ["Confiança", "Sofisticação", "Leveza", "Autenticidade", "Excelência em Experiência da Cliente"],
    universo: ["Praia", "Resorts", "Beach Clubs", "Viagens", "Verão"],
  },

  tom_de_voz: {
    personalidade: "Amiga sofisticada que entende de moda e lifestyle — elegante, próxima, consultiva e solar.",
    caracteristicas: ["Natural", "Consultiva", "Direta", "Elegante"],
    vocabulario_obrigatorio: [
      "Curadoria (ex: 'nossa curadoria de biquínis')",
      "Drop (substitui 'coleção' ou 'lançamento')",
      "Resort Wear (saídas de praia e vestuário para resorts)",
      "Beach Club (cenário de uso ideal)",
      "Must-have (peças essenciais)",
      "Fit Perfeito (caimento superior e modelagem exclusiva)",
    ],
    emojis: {
      "🖤": "Assinatura da marca — usar ao final das mensagens",
      "✨": "Destaque, recomendação, algo especial",
      "✈️": "Viagem, destino, getaway",
      "🌊": "Praia, mar, frescor, lazer",
      "🥂": "Celebração, momento especial, sofisticação",
    },
    instrucoes: "Usar emojis com moderação e intenção. Nunca usar em excesso.",
  },

  contato: {
    whatsapp: "(61) 9.9919-4999",
    whatsapp_link: "https://wa.me/5561999194999",
    email_sac: "sac@nakedsw.com.br",
    site: "https://www.nakedsw.com.br",
    plataforma_trocas: "https://nakedswimwear.troquefacil.com.br",
    horario_atendimento: {
      segunda_sexta: "09h às 17h",
      sabado_domingo: "07h às 14h",
    },
    meta_primeira_resposta: "Até 15 minutos",
  },

  promocoes_ativas: {
    cupons: [
      {
        codigo: "NKSW10",
        desconto: "10% OFF",
        condicao: "Primeira compra — qualquer produto",
        validade: "Verificar validade antes de oferecer",
      },
      {
        codigo: "PIX",
        desconto: "5% OFF adicional",
        condicao: "Pagamento via Pix em qualquer pedido",
        validade: "Sempre ativo",
      },
    ],
    frete_gratis: {
      valor_minimo: "R$ 1.200,00",
      observacao: "Alguns produtos premium já incluem frete grátis independente do valor",
    },
    aviso: "Sempre confirme a validade do cupom antes de oferecer. Em caso de dúvida, não prometa.",
  },

  tabela_medidas: {
    instrucoes_medicao: {
      busto: "Contorne a parte mais alta do busto sem apertar",
      baixo_busto: "Contorne onde o busto encontra a caixa torácica",
      cintura: "Contorne a parte mais fina da cintura sem apertar (altura do cotovelo)",
      cintura_baixa: "Contorne na altura do osso da bacia sem apertar",
      quadril: "Contorne a parte mais alta do bumbum sem apertar",
    },
    biquinis_e_maios_cm: {
      PP: { numeracao: "34-36", busto: "74-79", baixo_busto: "64-69", cintura: "58-63", cintura_baixa: "71-76", quadril: "85-90" },
      P:  { numeracao: "36-38", busto: "80-85", baixo_busto: "70-75", cintura: "64-69", cintura_baixa: "77-82", quadril: "91-96" },
      M:  { numeracao: "38-40", busto: "86-91", baixo_busto: "76-81", cintura: "70-75", cintura_baixa: "83-88", quadril: "97-102" },
      G:  { numeracao: "40-42", busto: "92-97", baixo_busto: "82-87", cintura: "76-81", cintura_baixa: "89-94", quadril: "103-108" },
      GG: { numeracao: "42-44", busto: "98-103", baixo_busto: "88-93", cintura: "82-87", cintura_baixa: "95-100", quadril: "109-114" },
    },
    roupas_cm: {
      PP: { numeracao: "34-36", busto: "80-84", cintura: "65-69", quadril: "86-90" },
      P:  { numeracao: "36-38", busto: "85-89", cintura: "70-74", quadril: "91-95" },
      M:  { numeracao: "38-40", busto: "90-94", cintura: "75-79", quadril: "96-100" },
      G:  { numeracao: "40-42", busto: "95-99", cintura: "80-84", quadril: "101-105" },
      GG: { numeracao: "42-44", busto: "100-104", cintura: "85-89", quadril: "106-110" },
    },
    dica: "Se a cliente estiver entre dois tamanhos, sugira o maior para mais conforto.",
  },

  politica_frete: {
    frete_gratis_acima: "R$ 1.200,00",
    transportadoras: ["Correios PAC", "Correios SEDEX", "Total Express"],
    como_calcular: "Cliente insere o CEP na página do carrinho para ver opções e valores",
    prazo_composicao: "Prazo de postagem + prazo de entrega da transportadora escolhida",
    atencao: [
      "Pedidos feitos sábados, domingos ou feriados: prazo inicia no próximo dia útil",
      "Produtos enviados em embalagem exclusiva NKSW com lacre de segurança",
      "Se o lacre chegar violado: não aceitar e contatar sac@nakedsw.com.br imediatamente",
    ],
    rastreamento: "Código enviado por e-mail após o despacho. Rastrear em: rastreamento.correios.com.br",
  },

  politica_trocas_devolucoes: {
    prazo: "7 dias corridos após o recebimento do produto",
    plataforma: "https://nakedswimwear.troquefacil.com.br",
    como_funciona: "O valor do produto vira crédito no site — não é troca direta por outro produto",
    custos: {
      primeira_troca: "NKSW arca com o frete de retorno. O envio do novo produto é por conta da cliente.",
      trocas_subsequentes: "Todos os custos (retorno e envio) são por conta da cliente.",
    },
    restricoes: [
      "NÃO aceitamos trocas de produtos em promoção ou do Outlet",
      "Produto deve estar sem uso, com etiqueta e embalagem original",
    ],
    script_troca:
      "Poxa, sinto muito que não tenha ficado como você esperava. ✨ Vamos resolver isso agora! Acessa nakedswimwear.troquefacil.com.br — é rápido e prático. Você recebe o código de postagem e depois o crédito para escolher a peça ideal. Me conta o que você sentiu falta, assim posso te ajudar a encontrar o fit perfeito na nossa curadoria. 🖤",
  },

  instrucoes_cuidado: {
    geral: [
      "Lavar à mão, separadamente por cor, com sabão neutro (pH abaixo de 7)",
      "Não lavar com água morna",
      "Enxaguar em água fria logo após o uso",
      "Secar à sombra, sem dobras, em ambiente arejado",
      "Não deixar secar no box do banheiro ou ambientes fechados",
      "Não torcer nem colocar na máquina de lavar",
    ],
    atencoes_por_cor: {
      cores_vivas: "Podem soltar tinta nos primeiros usos. Não deixar em contato com outras roupas quando molhadas.",
      cores_citricas: "Menos resistência à luz solar — pode haver alteração da cor com o tempo.",
      cores_claras: "Mais transparentes e sujeitas a manchas. Evitar contato com produtos corporais.",
    },
  },

  naked_atelier: {
    descricao:
      "Linha especial de peças feitas à mão, em edição limitada, com materiais nobres e acabamentos impecáveis. Cada peça é única — criada para quem valoriza o que é raro e especial.",
    diferenciais: ["Feito à mão do corte à costura final", "Edição limitada", "Materiais nobres selecionados", "Costuras invisíveis e acabamento de alto padrão"],
  },

  metodo_atendimento: {
    etapas: [
      { nome: "Recepção", objetivo: "Criar conexão imediata e personalizada", acao: "Cumprimentar pelo nome se disponível. Ex: 'Olá, Ana! Que bom ter você aqui. 🖤'" },
      { nome: "Diagnóstico", objetivo: "Entender o contexto completo", acao: "Perguntar: Onde será usada a peça? Qual o destino? Qual tipo de fit prefere?" },
      { nome: "Sugestão", objetivo: "Oferecer soluções precisas e curadas", acao: "Apresentar no máximo 3 opções, destacando tecido, modelagem e ocasião de uso." },
      { nome: "Fechamento", objetivo: "Facilitar a conversão", acao: "Enviar link direto do produto, informar prazo de entrega, tirar última dúvida." },
    ],
  },

  escalacao_humano: {
    quando_escalar: [
      "Reembolso em dinheiro (não vale-troca)",
      "Produto com defeito grave",
      "Cliente VIP (3+ compras) insatisfeita",
      "Pedido extraviado confirmado",
      "Atraso acima de 15 dias além do prometido",
      "Cliente pede explicitamente para falar com humano",
      "Reclamação repetida sem resolução",
      "Parcerias, influencers, B2B",
    ],
    script:
      "Entendo, vou conectar você agora com nossa equipe que vai conseguir resolver isso da melhor forma. Já passo todo o histórico da nossa conversa pra elas — você não vai precisar repetir nada. Um momento! 🖤",
    contexto_para_humano:
      "Ao transferir, enviar internamente: Nome do cliente, canal, motivo detalhado, número do pedido se houver, histórico da conversa, sentimento (satisfeita/insatisfeita/neutra), o que já foi oferecido e urgência (alta/média/baixa).",
  },

  faqs: [
    {
      pergunta: "Qual o prazo de entrega?",
      resposta:
        "O prazo é composto pelo tempo de postagem + prazo da transportadora escolhida, contados a partir da confirmação do pagamento. Pedidos feitos nos fins de semana ou feriados têm o prazo iniciado no próximo dia útil. Você pode calcular o frete e prazo exato inserindo seu CEP no carrinho. 🖤",
    },
    {
      pergunta: "Posso parcelar?",
      resposta:
        "Sim! Aceitamos parcelamento sem juros. Pagando via Pix você ainda ganha 5% de desconto adicional. ✨",
    },
    {
      pergunta: "Como sei meu tamanho?",
      resposta:
        "Me passa suas medidas de busto, cintura e quadril (em cm) que eu te indico o tamanho certinho! Se preferir, tem nossa tabela de medidas completa no site. Se estiver entre dois tamanhos, recomendamos o maior para mais conforto. 🖤",
    },
    {
      pergunta: "Vocês têm loja física?",
      resposta:
        "Somos uma marca digital — vendemos exclusivamente pelo site nakedsw.com.br. Isso nos permite oferecer o melhor custo-benefício e chegar a qualquer lugar do Brasil. ✨",
    },
    {
      pergunta: "Os produtos do Outlet têm troca?",
      resposta:
        "Produtos em promoção no Outlet não estão incluídos na nossa política de trocas. Todos os outros produtos do site têm troca facilitada em até 7 dias após o recebimento. 🖤",
    },
    {
      pergunta: "O que é o Naked Atelier?",
      resposta:
        "O Naked Atelier é nossa linha especial — peças feitas à mão, em edição limitada, com materiais nobres e acabamentos impecáveis. São peças únicas para quem valoriza o que é raro e especial. ✨🖤",
    },
  ],
};

export default function handler(req, res) {
  // CORS para qualquer origem (Voll pode chamar de qualquer lugar)
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "GET, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");

  if (req.method === "OPTIONS") {
    return res.status(200).end();
  }

  if (req.method !== "GET") {
    return res.status(405).json({ error: "Método não permitido" });
  }

  // Filtro por seção específica (ex: /api/knowledge?section=tabela_medidas)
  const { section } = req.query;
  if (section && knowledge[section]) {
    return res.status(200).json({ section, data: knowledge[section] });
  }

  return res.status(200).json(knowledge);
}
