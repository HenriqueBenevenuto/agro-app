/**
 * ===================================================================
 *  APP AVES VIVAS — BACKEND (Google Apps Script)
 * ===================================================================
 *  Como instalar:
 *  1. Abra a planilha "Planilha Venda Aves Vivas" no Google Sheets.
 *  2. Menu Extensões > Apps Script.
 *  3. Apague o conteúdo padrão do arquivo Code.gs e cole TODO este arquivo.
 *  4. Clique em Implantar (Deploy) > Nova implantação.
 *     - Tipo: App da Web
 *     - Executar como: Eu (seu e-mail)
 *     - Quem pode acessar: Qualquer pessoa
 *  5. Copie a URL do app da Web gerada (termina em /exec) e cole na
 *     tela de configuração do app (ícone de engrenagem).
 *  6. Sempre que editar este script, gere uma NOVA implantação
 *     (ou "Gerenciar implantações" > editar > Nova versão) para as
 *     mudanças valerem no link já publicado.
 * ===================================================================
 */

// Nome exato das abas de pedidos e como cada uma é organizada.
// "subtipoLabel" é o nome da coluna extra de variação (Cor, Tipo), ou null se não existir.
// Isso agora é editável pelo app (renomear tipo, criar tipo novo) — o que
// está aqui embaixo é só o ponto de partida (usado a primeira vez, antes
// de qualquer edição), guardado de fato em PropertiesService a partir daí.
const CATEGORIAS_AVES_PADRAO = [
  { nome: "Pinto de Corte Branco", subtipoLabel: null },
  { nome: "Postura 21 dias", subtipoLabel: "Cor" },
  { nome: "Machos 21 dias", subtipoLabel: "Tipo" },
  { nome: "Codorna", subtipoLabel: null },
  { nome: "Postura 90 dias", subtipoLabel: null }
];
const CATEGORIAS_AVES_KEY = "categoriasAvesRegistry";

function getCategoriasAvesRegistro() {
  const raw = PropertiesService.getScriptProperties().getProperty(CATEGORIAS_AVES_KEY);
  if (!raw) {
    PropertiesService.getScriptProperties().setProperty(CATEGORIAS_AVES_KEY, JSON.stringify(CATEGORIAS_AVES_PADRAO));
    return CATEGORIAS_AVES_PADRAO.slice();
  }
  try {
    const arr = JSON.parse(raw);
    return Array.isArray(arr) && arr.length ? arr : CATEGORIAS_AVES_PADRAO.slice();
  } catch (e) {
    return CATEGORIAS_AVES_PADRAO.slice();
  }
}

function salvarCategoriasAvesRegistro(lista) {
  PropertiesService.getScriptProperties().setProperty(CATEGORIAS_AVES_KEY, JSON.stringify(lista));
}

function categoriaAveExiste(categoria) {
  return getCategoriasAvesRegistro().some(function (c) { return c.nome === categoria; });
}

function getSubtipoLabel(categoria) {
  const item = getCategoriasAvesRegistro().filter(function (c) { return c.nome === categoria; })[0];
  return item ? (item.subtipoLabel || null) : null;
}

// Renomear só atualiza o cadastro e as linhas já lançadas na aba única de
// pedidos -- não precisa mais renomear nenhuma aba da planilha.
function renomearCategoriaAve(p) {
  const nomeAntigo = p.nomeAntigo;
  const nomeNovo = (p.nomeNovo || "").trim();
  if (!nomeNovo) throw new Error("Informe o novo nome.");
  const registro = getCategoriasAvesRegistro();
  const item = registro.filter(function (c) { return c.nome === nomeAntigo; })[0];
  if (!item) throw new Error("Tipo de ave não encontrado: " + nomeAntigo);
  if (nomeNovo !== nomeAntigo && registro.some(function (c) { return c.nome === nomeNovo; })) {
    throw new Error("Já existe um tipo de ave com esse nome.");
  }
  if (nomeNovo !== nomeAntigo) {
    item.nome = nomeNovo;
    salvarCategoriasAvesRegistro(registro);

    // Atualiza o nome em todos os pedidos já lançados com o nome antigo.
    const sheet = getAvesPedidoSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow >= AVES_PEDIDO_DATA_START_ROW) {
      const catCol = 5;
      const valores = sheet.getRange(AVES_PEDIDO_DATA_START_ROW, catCol, lastRow - AVES_PEDIDO_DATA_START_ROW + 1, 1).getValues();
      for (let i = 0; i < valores.length; i++) {
        if (valores[i][0] === nomeAntigo) {
          sheet.getRange(AVES_PEDIDO_DATA_START_ROW + i, catCol).setValue(nomeNovo);
        }
      }
    }

    // Migra os valores guardados fora da planilha (estoque, preço, extra
    // de encomenda) pra ficarem associados ao nome novo.
    const props = PropertiesService.getScriptProperties();
    [ESTOQUE_PREFIX, PRECO_PREFIX, ENCOMENDA_EXTRA_PREFIX].forEach(function (prefixo) {
      const valorAntigo = props.getProperty(prefixo + nomeAntigo);
      if (valorAntigo !== null) {
        props.setProperty(prefixo + nomeNovo, valorAntigo);
        props.deleteProperty(prefixo + nomeAntigo);
      }
    });
  }
  return { categorias: getCategoriasAvesRegistro() };
}

function criarCategoriaAve(p) {
  const nome = (p.nome || "").trim();
  if (!nome) throw new Error("Informe o nome do novo tipo de ave.");
  const subtipoLabel = (p.subtipoLabel || "").trim() || null;
  const registro = getCategoriasAvesRegistro();
  if (registro.some(function (c) { return c.nome === nome; })) {
    throw new Error("Já existe um tipo de ave com esse nome.");
  }
  registro.push({ nome: nome, subtipoLabel: subtipoLabel });
  salvarCategoriasAvesRegistro(registro);
  return { categorias: getCategoriasAvesRegistro() };
}

// Preços padrão por ave (usados na primeira vez, antes de qualquer edição no app).
const DEFAULT_PRECOS = {
  "Pinto de Corte Branco": 14.50,
  "Postura 21 dias": 22.60,
  "Machos 21 dias": 18.50,
  "Codorna": 22.60,
  "Postura 90 dias": 52.00
};

const CLIENTES_SHEET = "Clientes";
const CLIENTE_COLS = ["Nome", "Telefone", "Endereco", "Municipio"];

const HEADER_ROW = 3;   // linha onde ficam os títulos das colunas em todas as abas
const DATA_START_ROW = 4; // primeira linha de dados

// Senha exigida em toda chamada. Só existe aqui (nunca é enviada ao
// arquivo do app) — troque o texto abaixo quando quiser mudar a senha.
const APP_SENHA = "Agro3692";

// ---------------------------------------------------------------
// ENTRADA HTTP
// ---------------------------------------------------------------
function doGet(e) {
  return respond({ ok: true, msg: "API Aves Vivas no ar." });
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    if (String(body.senha || "") !== APP_SENHA) {
      return respond({ ok: false, error: "Senha incorreta." });
    }

    const action = body.action;
    const payload = body.payload || {};
    let result;

    switch (action) {
      case "listClientes": result = listClientes(); break;
      case "addCliente": result = addCliente(payload); break;
      case "updateCliente": result = updateCliente(payload); break;
      case "deleteCliente": result = deleteCliente(payload); break;

      case "listPedidos": result = listPedidos(); break;
      case "addPedido": result = addPedidoMulti(payload); break;
      case "updatePedido": result = updatePedidoMulti(payload); break;
      case "deletePedido": result = deletePedidoGroup(payload); break;
      case "encerrarPedido": result = setPedidoGroupStatus(payload, "Encerrado"); break;
      case "reabrirPedido": result = setPedidoGroupStatus(payload, "Ativo"); break;

      case "getMeta": result = getMeta(); break;
      case "renomearCategoriaAve": result = renomearCategoriaAve(payload); break;
      case "criarCategoriaAve": result = criarCategoriaAve(payload); break;
      case "getEstoque": result = getEstoque(); break;
      case "setEstoqueTotal": result = setEstoqueTotal(payload); break;
      case "getEncomendasExtras": result = getEncomendasExtras(); break;
      case "setEncomendaExtra": result = setEncomendaExtra(payload); break;
      case "getHistoricoEncomendas": result = getHistoricoEncomendas(); break;
      case "registrarEncomenda": result = registrarEncomenda(payload); break;
      case "confirmarEncomendaLote": result = confirmarEncomendaLote(payload); break;
      case "listPedidosPorLote": result = listPedidosPorLote(payload.loteId); break;
      case "getPendenciasLote": result = getPendenciasLote(payload.loteId); break;
      case "getGtaLoteInfo": result = getGtaLoteInfo(payload.loteId); break;
      case "getUltimoLoteId": result = getUltimoLoteId(); break;
      case "registrarRelacaoVenda": result = registrarRelacaoVenda(payload); break;
      case "listRelacoesVenda": result = listRelacoesVenda(); break;
      case "excluirRelacaoVenda": result = excluirRelacaoVenda(payload); break;
      case "venderAvulso": result = venderAvulso(payload); break;
      case "registrarRecebido": result = registrarRecebido(payload); break;

      case "listMaquinas": result = listMaquinas(); break;
      case "addMaquina": result = addMaquina(payload); break;
      case "updateMaquina": result = updateMaquina(payload); break;
      case "deleteMaquina": result = deleteMaquina(payload); break;

      case "listAlugueis": result = listAlugueis(); break;
      case "addAluguel": result = addAluguel(payload); break;
      case "updateAluguel": result = updateAluguel(payload); break;
      case "deleteAluguel": result = deleteAluguel(payload); break;
      case "encerrarAluguel": result = setAluguelStatus(payload, "Encerrado"); break;
      case "reabrirAluguel": result = setAluguelStatus(payload, "Ativo"); break;
      case "getPrecos": result = getPrecos(); break;
      case "setPreco": result = setPreco(payload); break;

      case "listPlantas": result = listPlantas(); break;
      case "addPlanta": result = addPlanta(payload); break;
      case "updatePlanta": result = updatePlanta(payload); break;
      case "deletePlanta": result = deletePlanta(payload); break;
      case "bulkAddPlantas": result = bulkAddPlantas(payload); break;

      case "listPedidosPlantas": result = listPedidosPlantas(); break;
      case "addPedidoPlanta": result = addPedidoPlantaMulti(payload); break;
      case "updatePedidoPlanta": result = updatePedidoPlantaMulti(payload); break;
      case "deletePedidoPlanta": result = deletePedidoPlantaGroup(payload); break;
      case "encerrarPedidoPlanta": result = setPedidoPlantaGroupStatus(payload, "Encerrado"); break;
      case "reabrirPedidoPlanta": result = setPedidoPlantaGroupStatus(payload, "Ativo"); break;

      case "getInicial": result = getInicial(); break;
      case "getInicialCore": result = getInicialCore(); break;
      case "getInicialResto": result = getInicialResto(); break;

      default: throw new Error("Ação desconhecida: " + action);
    }

    return respond({ ok: true, data: result });
  } catch (err) {
    return respond({ ok: false, error: String(err.message || err) });
  }
}

function respond(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function ss() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

// ---------------------------------------------------------------
// META (categorias disponíveis, para o app montar filtros/formulários)
// ---------------------------------------------------------------
function getMeta() {
  return { categorias: getCategoriasAvesRegistro() };
}

// Busca tudo que o app precisa pra abrir, numa chamada só (bem mais rápido
// do que fazer 4-5 requisições separadas: cada uma tem seu próprio custo
// de ida-e-volta até o Google).
function getInicial() {
  return {
    meta: getMeta(),
    clientes: listClientes(),
    pedidos: listPedidos(),
    estoque: getEstoque(),
    precos: getPrecos(),
    plantas: listPlantas(),
    pedidosPlantas: listPedidosPlantas(),
    encomendasExtras: getEncomendasExtras(),
    historicoEncomendas: getHistoricoEncomendas(),
    maquinas: listMaquinas(),
    alugueis: listAlugueis(),
    ultimoLoteId: getUltimoLoteId(),
    relacoesVenda: listRelacoesVenda()
  };
}

// Versão dividida em duas partes, pra abrir a tela mais rápido: a
// primeira só busca o essencial pra abrir em Aves + Clientes; o resto
// (Plantas, Máquinas, histórico de encomendas) busca depois, em segundo
// plano, sem travar a abertura do app.
function getInicialCore() {
  return {
    meta: getMeta(),
    clientes: listClientes(),
    pedidos: listPedidos(),
    estoque: getEstoque(),
    precos: getPrecos(),
    ultimoLoteId: getUltimoLoteId()
  };
}

function getInicialResto() {
  return {
    plantas: listPlantas(),
    pedidosPlantas: listPedidosPlantas(),
    encomendasExtras: getEncomendasExtras(),
    historicoEncomendas: getHistoricoEncomendas(),
    maquinas: listMaquinas(),
    alugueis: listAlugueis(),
    relacoesVenda: listRelacoesVenda()
  };
}

// ---------------------------------------------------------------
// ESTOQUE (total encomendado por tipo de ave, salvo fora das planilhas)
// ---------------------------------------------------------------
const ESTOQUE_PREFIX = "estoque_";

function getEstoque() {
  const todas = PropertiesService.getScriptProperties().getProperties();
  const out = {};
  getCategoriasAvesRegistro().forEach(function (c) {
    const v = todas[ESTOQUE_PREFIX + c.nome];
    out[c.nome] = v ? Number(v) : 0;
  });
  return out;
}

function setEstoqueTotal(p) {
  if (!categoriaAveExiste(p.categoria)) throw new Error("Categoria de ave inválida: " + p.categoria);
  const total = Number(p.total);
  if (isNaN(total) || total < 0) throw new Error("Total inválido.");
  PropertiesService.getScriptProperties().setProperty(ESTOQUE_PREFIX + p.categoria, String(total));
  return getEstoque();
}

// ---------------------------------------------------------------
// ENCOMENDAS (quantidade extra desejada por tipo de ave, além dos
// pedidos ativos — usada pra montar o pedido de compra ao fornecedor)
// ---------------------------------------------------------------
const ENCOMENDA_EXTRA_PREFIX = "encomendaExtra_";

function getEncomendasExtras() {
  const todas = PropertiesService.getScriptProperties().getProperties();
  const out = {};
  getCategoriasAvesRegistro().forEach(function (c) {
    const v = todas[ENCOMENDA_EXTRA_PREFIX + c.nome];
    out[c.nome] = v ? Number(v) : 0;
  });
  return out;
}

function setEncomendaExtra(p) {
  if (!categoriaAveExiste(p.categoria)) throw new Error("Categoria de ave inválida: " + p.categoria);
  const extra = Number(p.extra);
  if (isNaN(extra) || extra < 0) throw new Error("Quantidade inválida.");
  PropertiesService.getScriptProperties().setProperty(ENCOMENDA_EXTRA_PREFIX + p.categoria, String(extra));
  return getEncomendasExtras();
}

// Guarda um histórico das últimas encomendas confirmadas (até 10), pra
// mostrar no app.
const HISTORICO_ENCOMENDAS_KEY = "historicoEncomendas";

function getHistoricoEncomendas() {
  const raw = PropertiesService.getScriptProperties().getProperty(HISTORICO_ENCOMENDAS_KEY);
  if (!raw) return [];
  try {
    const arr = JSON.parse(raw);
    return Array.isArray(arr) ? arr : [];
  } catch (e) {
    return [];
  }
}

function registrarEncomenda(p) {
  const registro = {
    data: Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT-3", "dd/MM/yyyy HH:mm"),
    itens: (p.itens || []).map(function (item) {
      return { categoria: item.categoria, qtd: Number(item.qtd) || 0 };
    })
  };
  const historico = getHistoricoEncomendas();
  historico.unshift(registro);
  const limitado = historico.slice(0, 10);
  PropertiesService.getScriptProperties().setProperty(HISTORICO_ENCOMENDAS_KEY, JSON.stringify(limitado));
  return limitado;
}

// Confirma a encomenda inteira numa chamada só: atualiza o estoque e zera
// o extra de cada tipo, e registra no histórico — tudo de uma vez, bem
// mais rápido do que fazer isso categoria por categoria.
function confirmarEncomendaLote(p) {
  const itens = p.itens || [];
  const props = PropertiesService.getScriptProperties();
  itens.forEach(function (item) {
    if (!categoriaAveExiste(item.categoria)) return;
    const total = Number(item.total) || 0;
    props.setProperty(ESTOQUE_PREFIX + item.categoria, String(total));
    props.setProperty(ENCOMENDA_EXTRA_PREFIX + item.categoria, "0");
  });
  const historico = registrarEncomenda({
    itens: itens.map(function (item) { return { categoria: item.categoria, qtd: Number(item.total) || 0 }; })
  });

  // Gera um Lote ID novo e marca todos os pedidos ativos desses tipos com
  // ele, pra depois dar pra imprimir a relação de venda (com GTA) desse
  // lote especificamente.
  const loteId = Utilities.getUuid();
  marcarLoteParaCategorias(itens.map(function (item) { return item.categoria; }), loteId);

  return {
    estoque: getEstoque(),
    encomendasExtras: getEncomendasExtras(),
    historico: historico,
    loteId: loteId
  };
}

// Registra quantas aves foram realmente recebidas do fornecedor pra um
// tipo de ave, dentro da encomenda mais recente do histórico.
function registrarRecebido(p) {
  const historico = getHistoricoEncomendas();
  if (!historico.length) throw new Error("Nenhuma encomenda confirmada ainda.");
  const recebido = Number(p.recebido);
  if (isNaN(recebido) || recebido < 0) throw new Error("Quantidade inválida.");
  const item = historico[0].itens.filter(function (it) { return it.categoria === p.categoria; })[0];
  if (!item) throw new Error("Esse tipo não faz parte da última encomenda.");
  item.recebido = recebido;
  PropertiesService.getScriptProperties().setProperty(HISTORICO_ENCOMENDAS_KEY, JSON.stringify(historico));
  return historico;
}

// ---------------------------------------------------------------
// PREÇOS (valor por ave, salvo fora das planilhas)
// ---------------------------------------------------------------
const PRECO_PREFIX = "preco_";

function getPrecos() {
  const todas = PropertiesService.getScriptProperties().getProperties();
  const out = {};
  getCategoriasAvesRegistro().forEach(function (c) {
    const v = todas[PRECO_PREFIX + c.nome];
    out[c.nome] = v ? Number(v) : (DEFAULT_PRECOS[c.nome] || 0);
  });
  return out;
}

function setPreco(p) {
  if (!categoriaAveExiste(p.categoria)) throw new Error("Categoria de ave inválida: " + p.categoria);
  const preco = Number(p.preco);
  if (isNaN(preco) || preco < 0) throw new Error("Preço inválido.");
  PropertiesService.getScriptProperties().setProperty(PRECO_PREFIX + p.categoria, String(preco));
  return getPrecos();
}

// ---------------------------------------------------------------
// CLIENTES
// ---------------------------------------------------------------
function getClientesSheet() {
  const sheet = ss().getSheetByName(CLIENTES_SHEET);
  if (!sheet) throw new Error("Aba 'Clientes' não encontrada.");
  return sheet;
}

function listClientes() {
  const sheet = getClientesSheet();
  const maxRow = sheet.getLastRow();
  if (maxRow < DATA_START_ROW) return [];
  const values = sheet.getRange(DATA_START_ROW, 1, maxRow - DATA_START_ROW + 1, 4).getValues();
  const out = [];
  values.forEach(function (row, i) {
    if (!row[0] || isPlaceholderRow(row[0])) return; // sem nome ou linha de exemplo, pula
    out.push({
      rowIndex: DATA_START_ROW + i,
      nome: row[0],
      telefone: normalizeStr(row[1]),
      endereco: normalizeStr(row[2]),
      municipio: normalizeStr(row[3])
    });
  });
  return out;
}

function addCliente(p) {
  const sheet = getClientesSheet();
  const lastRow = lastNonEmptyRow(sheet, 1, DATA_START_ROW);
  const newRow = Math.max(lastRow + 1, DATA_START_ROW);
  sheet.getRange(newRow, 1, 1, 4).setValues([[
    p.nome || "", p.telefone || "", p.endereco || "", p.municipio || ""
  ]]);
  return { rowIndex: newRow };
}

function updateCliente(p) {
  const sheet = getClientesSheet();
  requireRow(p.rowIndex);
  sheet.getRange(p.rowIndex, 1, 1, 4).setValues([[
    p.nome || "", p.telefone || "", p.endereco || "", p.municipio || ""
  ]]);
  return { rowIndex: p.rowIndex };
}

function deleteCliente(p) {
  const sheet = getClientesSheet();
  requireRow(p.rowIndex);
  sheet.deleteRow(p.rowIndex);
  return { deleted: true };
}

// ---------------------------------------------------------------
// PEDIDOS
// Todos os pedidos de ave ficam numa aba única ("Pedidos Aves"), com o
// tipo de ave numa coluna ("Categoria") em vez de uma aba por tipo.
// Cada pedido pode ter vários itens; todas as linhas de um mesmo pedido
// compartilham um "Pedido ID", gravado numa coluna extra.
// ---------------------------------------------------------------
const AVES_PEDIDO_SHEET = "Pedidos Aves";
const AVES_PEDIDO_HEADER_ROW = 1;
const AVES_PEDIDO_DATA_START_ROW = 2;
const AVES_PEDIDO_COLS = ["Nome", "Telefone", "Endereco", "Municipio", "Categoria", "Subtipo", "Qtd", "Observacoes"];
// Colunas depois das de cima: Status, Data Criacao, Pedido ID, Lote ID, Valor Pago.
const AVES_PEDIDO_TOTAL_COLS = AVES_PEDIDO_COLS.length + 5;
const AVES_PEDIDO_VALOR_PAGO_COL = AVES_PEDIDO_COLS.length + 5; // coluna 13 (M)

function getAvesPedidoSheet() {
  let sheet = ss().getSheetByName(AVES_PEDIDO_SHEET);
  if (!sheet) {
    sheet = ss().insertSheet(AVES_PEDIDO_SHEET);
    const headers = AVES_PEDIDO_COLS.concat(["Status", "Data Criacao", "Pedido ID", "Lote ID", "Valor Pago"]);
    sheet.getRange(AVES_PEDIDO_HEADER_ROW, 1, 1, headers.length).setValues([headers]);
  }
  garantirColunaValorPago(sheet);
  return sheet;
}

// A coluna "Valor Pago" foi adicionada depois -- em planilhas antigas ela
// ainda não existe. Cria o título (e a coluna, se a aba for estreita) uma
// vez só; nas próximas chamadas não faz nada.
let colunaValorPagoConferida = false;
function garantirColunaValorPago(sheet) {
  if (colunaValorPagoConferida) return;
  if (sheet.getMaxColumns() < AVES_PEDIDO_VALOR_PAGO_COL) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), AVES_PEDIDO_VALOR_PAGO_COL - sheet.getMaxColumns());
  }
  const titulo = sheet.getRange(AVES_PEDIDO_HEADER_ROW, AVES_PEDIDO_VALOR_PAGO_COL);
  if (!titulo.getValue()) titulo.setValue("Valor Pago");
  colunaValorPagoConferida = true;
}

function listPedidos() {
  const sheet = getAvesPedidoSheet();
  const totalCols = AVES_PEDIDO_TOTAL_COLS; // + Status + Data + Pedido ID + Lote ID + Valor Pago
  const lastRow = sheet.getLastRow();
  if (lastRow < AVES_PEDIDO_DATA_START_ROW) return [];
  const values = sheet.getRange(AVES_PEDIDO_DATA_START_ROW, 1, lastRow - AVES_PEDIDO_DATA_START_ROW + 1, totalCols).getValues();

  const subtipoLabelPorCategoria = {};
  getCategoriasAvesRegistro().forEach(function (c) { subtipoLabelPorCategoria[c.nome] = c.subtipoLabel || null; });

  const all = [];
  values.forEach(function (row, i) {
    if (!row[0] || isPlaceholderRow(row[0])) return; // sem nome ou linha de exemplo, pula
    const rowIndex = AVES_PEDIDO_DATA_START_ROW + i;
    const categoria = row[4];
    const pedidoIdRaw = row[10];
    all.push({
      rowIndex: rowIndex,
      categoria: categoria,
      subtipoLabel: subtipoLabelPorCategoria[categoria] || null,
      nome: row[0],
      telefone: normalizeStr(row[1]),
      endereco: normalizeStr(row[2]),
      municipio: normalizeStr(row[3]),
      subtipo: normalizeStr(row[5]),
      qtd: row[6] || 0,
      observacoes: normalizeStr(row[7]),
      status: row[8] || "Ativo",
      dataCriacao: row[9] ? formatDate(row[9]) : "",
      pedidoId: pedidoIdRaw ? String(pedidoIdRaw) : (categoria + "#" + rowIndex),
      loteId: row[11] ? String(row[11]) : "",
      valorPago: Number(row[12]) || 0
    });
  });
  return all;
}

function buildAvesPedidoRow(p) {
  return [p.nome || "", p.telefone || "", p.endereco || "", p.municipio || "", p.categoria || "", p.subtipo || "", p.qtd || 0, p.observacoes || ""];
}

// Escreve uma linha por item, todas marcadas com o mesmo Pedido ID.
// Se "status" for informado (ex: venda avulsa), usa esse em vez de "Ativo".
function addPedidoMulti(p) {
  if (!p.itens || !p.itens.length) throw new Error("Informe ao menos um tipo de ave no pedido.");
  const pedidoId = p.pedidoId || Utilities.getUuid();
  let dataCriacao = new Date();
  // Quando o app separa o restante de um pedido (o fornecedor não mandou
  // todos os tipos), o pedido novo mantém a data do original, pra não
  // perder a ordem de quem pediu primeiro.
  if (p.copiarDataDoPedidoId) {
    const origem = findPedidoGroupRows(String(p.copiarDataDoPedidoId));
    if (origem.length && origem[0].dataCriacaoRaw) dataCriacao = origem[0].dataCriacaoRaw;
  }
  const status = p.status || "Ativo";
  const loteId = p.loteId || "";
  const valorPago = Math.max(0, Number(p.valorPago) || 0);
  const sheet = getAvesPedidoSheet();
  const criados = [];
  p.itens.forEach(function (item) {
    const newRow = Math.max(sheet.getLastRow() + 1, AVES_PEDIDO_DATA_START_ROW);
    const rowData = buildAvesPedidoRow({
      nome: p.nome, telefone: p.telefone, endereco: p.endereco, municipio: p.municipio,
      categoria: item.categoria, subtipo: item.subtipo, qtd: item.qtd, observacoes: p.observacoes
    });
    sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
    sheet.getRange(newRow, rowData.length + 1, 1, 5).setValues([[status, dataCriacao, pedidoId, loteId, valorPago]]);
    criados.push({ categoria: item.categoria, rowIndex: newRow });
  });
  return { pedidoId: pedidoId, itens: criados };
}

// Edita um pedido: apaga as linhas antigas do grupo e recria com os itens novos,
// preservando o Pedido ID, a data de criação original, o status e o Lote ID.
// O valor pago vem do app; se não vier (versão antiga do app), mantém o que já estava.
function updatePedidoMulti(p) {
  if (!p.pedidoId) throw new Error("Pedido inválido.");
  if (!p.itens || !p.itens.length) throw new Error("Informe ao menos um tipo de ave no pedido.");

  const existentes = findPedidoGroupRows(p.pedidoId);
  const dataOriginal = existentes.length && existentes[0].dataCriacaoRaw ? existentes[0].dataCriacaoRaw : new Date();
  const statusOriginal = existentes.length ? existentes[0].status : "Ativo";
  const loteOriginal = existentes.length ? (existentes[0].loteId || "") : "";
  const valorPagoOriginal = existentes.length ? (Number(existentes[0].valorPago) || 0) : 0;
  const valorPago = (p.valorPago === undefined || p.valorPago === null || p.valorPago === "")
    ? valorPagoOriginal
    : Math.max(0, Number(p.valorPago) || 0);

  deletePedidoGroup({ pedidoId: p.pedidoId });

  const sheet = getAvesPedidoSheet();
  const criados = [];
  p.itens.forEach(function (item) {
    const newRow = Math.max(sheet.getLastRow() + 1, AVES_PEDIDO_DATA_START_ROW);
    const rowData = buildAvesPedidoRow({
      nome: p.nome, telefone: p.telefone, endereco: p.endereco, municipio: p.municipio,
      categoria: item.categoria, subtipo: item.subtipo, qtd: item.qtd, observacoes: p.observacoes
    });
    sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
    sheet.getRange(newRow, rowData.length + 1, 1, 5).setValues([[statusOriginal, dataOriginal, p.pedidoId, loteOriginal, valorPago]]);
    criados.push({ categoria: item.categoria, rowIndex: newRow });
  });
  return { pedidoId: p.pedidoId, itens: criados };
}

// Acha todas as linhas que pertencem a um Pedido ID.
function findPedidoGroupRows(pedidoId) {
  const sheet = getAvesPedidoSheet();
  const totalCols = AVES_PEDIDO_TOTAL_COLS;
  const lastRow = sheet.getLastRow();
  const out = [];
  if (lastRow < AVES_PEDIDO_DATA_START_ROW) return out;
  const values = sheet.getRange(AVES_PEDIDO_DATA_START_ROW, 1, lastRow - AVES_PEDIDO_DATA_START_ROW + 1, totalCols).getValues();
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (!row[0]) continue;
    const rowIndex = AVES_PEDIDO_DATA_START_ROW + i;
    const categoria = row[4];
    const rowId = row[10] ? String(row[10]) : (categoria + "#" + rowIndex);
    if (rowId === pedidoId) {
      out.push({
        categoria: categoria,
        rowIndex: rowIndex,
        qtd: row[6] || 0,
        status: row[8] || "Ativo",
        dataCriacaoRaw: row[9],
        loteId: row[11] ? String(row[11]) : "",
        valorPago: Number(row[12]) || 0
      });
    }
  }
  return out;
}

// Remove o pedido de vez. Se alguma linha já estava Encerrada (as aves já
// tinham sido descontadas do estoque), devolve essa quantidade primeiro --
// senão o estoque ficaria descontado pra sempre por um pedido que nem
// existe mais.
function deletePedidoGroup(p) {
  if (!p.pedidoId) throw new Error("Pedido inválido.");
  const sheet = getAvesPedidoSheet();
  const rows = findPedidoGroupRows(p.pedidoId);

  const props = PropertiesService.getScriptProperties();
  rows.forEach(function (r) {
    if ((r.status || "Ativo") === "Encerrado") {
      const atual = Number(props.getProperty(ESTOQUE_PREFIX + r.categoria)) || 0;
      props.setProperty(ESTOQUE_PREFIX + r.categoria, String(atual + (Number(r.qtd) || 0)));
    }
  });

  const indices = rows.map(function (r) { return r.rowIndex; }).sort(function (a, b) { return b - a; });
  indices.forEach(function (rowIndex) { sheet.deleteRow(rowIndex); });
  return { deleted: true, count: rows.length };
}

// Encerrar um pedido significa que as aves saíram de fato -- por isso
// desconta do Estoque (Encomendado), igual já acontecia na venda avulsa.
// Reabrir devolve pro estoque (voltam a ficar reservadas, não entregues).
function setPedidoGroupStatus(p, status) {
  if (!p.pedidoId) throw new Error("Pedido inválido.");
  const sheet = getAvesPedidoSheet();
  const statusCol = AVES_PEDIDO_COLS.length + 1;
  const rows = findPedidoGroupRows(p.pedidoId);
  const props = PropertiesService.getScriptProperties();

  if (status === "Encerrado") {
    // Não deixa encerrar se não tiver ave suficiente sobrando em estoque
    // pra cobrir esse pedido.
    const necessarioPorCategoria = {};
    rows.forEach(function (r) {
      if ((r.status || "Ativo") !== "Ativo") return; // só o que vai mudar de fato
      necessarioPorCategoria[r.categoria] = (necessarioPorCategoria[r.categoria] || 0) + (Number(r.qtd) || 0);
    });
    const faltando = [];
    Object.keys(necessarioPorCategoria).forEach(function (categoria) {
      const disponivel = Number(props.getProperty(ESTOQUE_PREFIX + categoria)) || 0;
      if (necessarioPorCategoria[categoria] > disponivel) {
        faltando.push(categoria + " (precisa " + necessarioPorCategoria[categoria] + ", tem " + disponivel + " sobrando)");
      }
    });
    if (faltando.length > 0) {
      throw new Error("Não tem ave suficiente em estoque pra encerrar: " + faltando.join("; ") + ".");
    }
  }

  rows.forEach(function (r) {
    const statusAtual = r.status || "Ativo";
    if (statusAtual !== status) {
      const atual = Number(props.getProperty(ESTOQUE_PREFIX + r.categoria)) || 0;
      if (status === "Encerrado" && statusAtual === "Ativo") {
        props.setProperty(ESTOQUE_PREFIX + r.categoria, String(Math.max(0, atual - (Number(r.qtd) || 0))));
      } else if (status === "Ativo" && statusAtual === "Encerrado") {
        props.setProperty(ESTOQUE_PREFIX + r.categoria, String(atual + (Number(r.qtd) || 0)));
      }
    }
    sheet.getRange(r.rowIndex, statusCol).setValue(status);
  });
  return { status: status, count: rows.length };
}

// ---------------------------------------------------------------
// MIGRAÇÃO (rode UMA VEZ manualmente, aqui no editor, pra mover os
// pedidos das abas antigas -- uma por tipo de ave -- pra essa aba única
// nova. As abas antigas NÃO são apagadas automaticamente; confira se
// ficou tudo certo antes de apagá-las manualmente, se quiser.)
// ---------------------------------------------------------------
function migrarAvesParaAbaUnica() {
  const registro = getCategoriasAvesRegistro();
  const novaSheet = getAvesPedidoSheet();
  let totalMigrado = 0;

  registro.forEach(function (c) {
    const sheetAntiga = ss().getSheetByName(c.nome);
    if (!sheetAntiga) return; // não existe mais essa aba (ex: tipo criado depois)

    const cols = ["Nome", "Telefone", "Endereco", "Municipio"];
    if (c.subtipoLabel) cols.push(c.subtipoLabel);
    cols.push("Qtd", "Observacoes");
    const totalCols = cols.length + 4; // + Status, Data, Pedido ID, Lote ID

    const lastRow = sheetAntiga.getLastRow();
    if (lastRow < DATA_START_ROW) return;
    const values = sheetAntiga.getRange(DATA_START_ROW, 1, lastRow - DATA_START_ROW + 1, totalCols).getValues();

    const linhasNovas = [];
    values.forEach(function (row) {
      if (!row[0] || isPlaceholderRow(row[0])) return;
      let idx = 4;
      let subtipo = "";
      if (c.subtipoLabel) { subtipo = row[idx] || ""; idx++; }
      const qtd = row[idx] || 0; idx++;
      const observacoes = row[idx] || ""; idx++;
      const status = row[idx] || "Ativo"; idx++;
      const dataCriacao = row[idx] || ""; idx++;
      const pedidoId = row[idx] || ""; idx++;
      const loteId = row[idx] || "";
      linhasNovas.push([
        row[0], row[1] || "", row[2] || "", row[3] || "", c.nome, subtipo, qtd, observacoes,
        status, dataCriacao, pedidoId, loteId
      ]);
    });

    if (linhasNovas.length > 0) {
      const startRow = Math.max(novaSheet.getLastRow() + 1, AVES_PEDIDO_DATA_START_ROW);
      novaSheet.getRange(startRow, 1, linhasNovas.length, 12).setValues(linhasNovas);
      totalMigrado += linhasNovas.length;
    }
  });

  Logger.log("Migração concluída! " + totalMigrado + " pedidos movidos pra aba 'Pedidos Aves'.");
  Logger.log("As abas antigas (uma por tipo de ave) NÃO foram apagadas -- confira se está tudo certo no app antes de apagá-las manualmente, se quiser.");
}

// ---------------------------------------------------------------
// LOTES (encomenda confirmada) + RELAÇÕES DE VENDA (GTA)
// ---------------------------------------------------------------
const ULTIMO_LOTE_KEY = "ultimoLoteId";
const RELACOES_SHEET = "Relacoes de Venda";
const RELACOES_HEADER_ROW = 1;
const RELACOES_DATA_START_ROW = 2;

// Marca (com um novo Lote ID) todas as linhas ATIVAS das categorias
// informadas -- chamado ao confirmar uma encomenda, pra depois dar pra
// imprimir a relação de venda só desse lote.
function marcarLoteParaCategorias(categorias, loteId) {
  const sheet = getAvesPedidoSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow >= AVES_PEDIDO_DATA_START_ROW) {
    const totalCols = AVES_PEDIDO_COLS.length + 4;
    const values = sheet.getRange(AVES_PEDIDO_DATA_START_ROW, 1, lastRow - AVES_PEDIDO_DATA_START_ROW + 1, totalCols).getValues();
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      if (!row[0]) continue;
      const categoria = row[4];
      const status = row[8] || "Ativo";
      if (categorias.indexOf(categoria) !== -1 && status === "Ativo") {
        sheet.getRange(AVES_PEDIDO_DATA_START_ROW + i, 12).setValue(loteId);
      }
    }
  }
  PropertiesService.getScriptProperties().setProperty(ULTIMO_LOTE_KEY, loteId);
}

function getUltimoLoteId() {
  return PropertiesService.getScriptProperties().getProperty(ULTIMO_LOTE_KEY) || "";
}

// Itens de um Lote ID que já foram vendidos de fato (Encerrado) --
// usado pra montar a relação de venda pra imprimir.
function listPedidosPorLote(loteId) {
  if (!loteId) return [];
  return listPedidos().filter(function (p) { return p.loteId === loteId && p.status === "Encerrado"; });
}

// Verifica se ainda tem pedido em aberto ou ave sobrando desse lote --
// usado pra travar a impressão da GTA até o lote estar 100% resolvido.
function getPendenciasLote(loteId) {
  if (!loteId) return { temAtivos: false, categoriasComSobra: [] };
  const todosPedidos = listPedidos();
  const categoriasEnvolvidas = {};
  let temAtivos = false;

  todosPedidos.forEach(function (p) {
    if (p.loteId !== loteId) return;
    categoriasEnvolvidas[p.categoria] = true;
    if ((p.status || "Ativo") === "Ativo") temAtivos = true;
  });

  const estoqueAtual = getEstoque();
  const categoriasComSobra = Object.keys(categoriasEnvolvidas).map(function (categoria) {
    const ativos = todosPedidos
      .filter(function (p) { return p.categoria === categoria && (p.status || "Ativo") === "Ativo"; })
      .reduce(function (s, p) { return s + (Number(p.qtd) || 0); }, 0);
    const sobrando = (Number(estoqueAtual[categoria]) || 0) - ativos;
    return { categoria: categoria, sobrando: sobrando };
  }).filter(function (c) { return c.sobrando > 0; });

  return { temAtivos: temAtivos, categoriasComSobra: categoriasComSobra };
}

// Junta o que a aba GTA precisa numa chamada só (uma varredura das abas
// de ave em vez de duas) -- os pedidos concluídos do lote + as pendências.
function getGtaLoteInfo(loteId) {
  if (!loteId) return { itensConcluidos: [], temAtivos: false, categoriasComSobra: [] };
  const todosPedidos = listPedidos();
  const doLote = todosPedidos.filter(function (p) { return p.loteId === loteId; });

  const categoriasEnvolvidas = {};
  let temAtivos = false;
  const itensConcluidos = [];

  doLote.forEach(function (p) {
    categoriasEnvolvidas[p.categoria] = true;
    if ((p.status || "Ativo") === "Ativo") {
      temAtivos = true;
    } else {
      itensConcluidos.push({
        rowIndex: p.rowIndex, categoria: p.categoria, subtipo: p.subtipo, qtd: p.qtd, status: p.status,
        nome: p.nome, telefone: p.telefone, endereco: p.endereco, municipio: p.municipio, pedidoId: p.pedidoId
      });
    }
  });

  const estoqueAtual = getEstoque();
  const categoriasComSobra = Object.keys(categoriasEnvolvidas).map(function (categoria) {
    const ativos = todosPedidos
      .filter(function (p) { return p.categoria === categoria && (p.status || "Ativo") === "Ativo"; })
      .reduce(function (s, p) { return s + (Number(p.qtd) || 0); }, 0);
    const sobrando = (Number(estoqueAtual[categoria]) || 0) - ativos;
    return { categoria: categoria, sobrando: sobrando };
  }).filter(function (c) { return c.sobrando > 0; });

  return { itensConcluidos: itensConcluidos, temAtivos: temAtivos, categoriasComSobra: categoriasComSobra };
}

function getRelacoesSheet() {
  let sheet = ss().getSheetByName(RELACOES_SHEET);
  if (!sheet) {
    sheet = ss().insertSheet(RELACOES_SHEET);
    sheet.getRange(RELACOES_HEADER_ROW, 1, 1, 5).setValues([["GTA", "Data", "LoteId", "ItensJson", "ValorTotal"]]);
  }
  return sheet;
}

function registrarRelacaoVenda(p) {
  if (!p.gta) throw new Error("Informe o código GTA.");
  if (!p.itens || !p.itens.length) throw new Error("Nenhum item pra registrar.");
  if (p.loteId) {
    const pendencias = getPendenciasLote(p.loteId);
    if (pendencias.temAtivos) throw new Error("Ainda tem pedido em aberto nesse lote. Encerre ou remova todos antes de imprimir a GTA.");
    if (pendencias.categoriasComSobra.length > 0) throw new Error("Ainda tem ave sobrando desse lote (" + pendencias.categoriasComSobra.map(function (c) { return c.categoria; }).join(", ") + "). Confirme a venda de tudo antes de imprimir a GTA.");
  }
  const sheet = getRelacoesSheet();
  const newRow = Math.max(sheet.getLastRow() + 1, RELACOES_DATA_START_ROW);
  const valorTotal = p.itens.reduce(function (s, it) { return s + (Number(it.valorTotal) || 0); }, 0);
  sheet.getRange(newRow, 1, 1, 5).setValues([[
    p.gta, new Date(), p.loteId || "", JSON.stringify(p.itens), valorTotal
  ]]);

  // "Consome" os pedidos que entraram nessa relação: limpa o Lote ID deles
  // pra não aparecerem de novo como pendentes na aba GTA.
  const loteCol = AVES_PEDIDO_COLS.length + 4;
  p.itens.forEach(function (item) {
    if (!item.rowIndex) return;
    try {
      getAvesPedidoSheet().getRange(item.rowIndex, loteCol).setValue("");
    } catch (e) {
      // se a linha nao existir mais (foi removida nesse meio tempo), ignora
    }
  });

  return listRelacoesVenda();
}

function listRelacoesVenda() {
  const sheet = getRelacoesSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < RELACOES_DATA_START_ROW) return [];
  const values = sheet.getRange(RELACOES_DATA_START_ROW, 1, lastRow - RELACOES_DATA_START_ROW + 1, 5).getValues();
  const out = [];
  values.forEach(function (row, i) {
    if (!row[0]) return;
    let itens = [];
    try { itens = JSON.parse(row[3] || "[]"); } catch (e) { itens = []; }
    out.push({
      rowIndex: RELACOES_DATA_START_ROW + i,
      gta: String(row[0]),
      data: row[1] ? formatDate(row[1]) : "",
      loteId: row[2] || "",
      itens: itens,
      valorTotal: Number(row[4]) || 0
    });
  });
  return out.reverse(); // mais recente primeiro
}

function excluirRelacaoVenda(p) {
  if (!p.rowIndex) throw new Error("Linha inválida.");
  getRelacoesSheet().deleteRow(p.rowIndex);
  return listRelacoesVenda();
}

// Venda avulsa: registra uma venda direto como Encerrado (sem passar por
// "Ativo"), marcada com o Lote ID informado, e desconta do Estoque na hora
// (já que representa aves que saíram de verdade, fora de qualquer pedido).
function venderAvulso(p) {
  if (!p.itens || !p.itens.length) throw new Error("Informe ao menos um tipo de ave.");
  const resultado = addPedidoMulti({
    nome: p.nome, telefone: p.telefone, endereco: p.endereco, municipio: p.municipio,
    observacoes: p.observacoes || "Venda avulsa", itens: p.itens,
    status: "Encerrado", loteId: p.loteId || ""
  });

  // Desconta do estoque (Encomendado) o que foi vendido, pra refletir no Sobrando.
  const props = PropertiesService.getScriptProperties();
  const todas = props.getProperties();
  p.itens.forEach(function (item) {
    const atual = Number(todas[ESTOQUE_PREFIX + item.categoria]) || 0;
    const novo = Math.max(0, atual - (Number(item.qtd) || 0));
    props.setProperty(ESTOQUE_PREFIX + item.categoria, String(novo));
  });

  return resultado;
}

// ---------------------------------------------------------------
// PLANTAS — catálogo (Cadastro) e pedidos
// Diferente das aves, os tipos de planta não são fixos: o usuário
// cadastra/edita/remove pelo próprio app. Por isso os pedidos de
// plantas ficam todos numa única aba (não uma aba por tipo).
// ---------------------------------------------------------------
const PLANTA_HEADER_ROW = 1;
const PLANTA_DATA_START_ROW = 2;
const PLANTAS_SHEET = "Plantas";
const PEDIDOS_PLANTAS_SHEET = "Pedidos Plantas";
const PLANTA_PEDIDO_COLS = ["Nome", "Telefone", "Endereco", "Municipio", "Planta", "Qtd", "Observacoes"];

function getPlantasSheet() {
  let sheet = ss().getSheetByName(PLANTAS_SHEET);
  if (!sheet) {
    sheet = ss().insertSheet(PLANTAS_SHEET);
    sheet.getRange(PLANTA_HEADER_ROW, 1, 1, 3).setValues([["Nome", "Preco", "Categoria"]]);
  }
  return sheet;
}

function listPlantas() {
  const sheet = getPlantasSheet();
  const maxRow = sheet.getLastRow();
  if (maxRow < PLANTA_DATA_START_ROW) return [];
  const values = sheet.getRange(PLANTA_DATA_START_ROW, 1, maxRow - PLANTA_DATA_START_ROW + 1, 3).getValues();
  const out = [];
  values.forEach(function (row, i) {
    if (!row[0] || isPlaceholderRow(row[0])) return;
    out.push({ rowIndex: PLANTA_DATA_START_ROW + i, nome: row[0], preco: Number(row[1]) || 0, categoria: normalizeStr(row[2]) });
  });
  return out;
}

function addPlanta(p) {
  const sheet = getPlantasSheet();
  const newRow = Math.max(sheet.getLastRow() + 1, PLANTA_DATA_START_ROW);
  sheet.getRange(newRow, 1, 1, 3).setValues([[p.nome || "", Number(p.preco) || 0, p.categoria || ""]]);
  return { rowIndex: newRow };
}

function updatePlanta(p) {
  const sheet = getPlantasSheet();
  if (!p.rowIndex || p.rowIndex < PLANTA_DATA_START_ROW) throw new Error("Linha inválida.");
  sheet.getRange(p.rowIndex, 1, 1, 3).setValues([[p.nome || "", Number(p.preco) || 0, p.categoria || ""]]);
  return { rowIndex: p.rowIndex };
}

function deletePlanta(p) {
  const sheet = getPlantasSheet();
  if (!p.rowIndex || p.rowIndex < PLANTA_DATA_START_ROW) throw new Error("Linha inválida.");
  sheet.deleteRow(p.rowIndex);
  return { deleted: true };
}

// Cadastra várias plantas de uma vez só (uma única escrita na planilha,
// bem mais rápido que chamar addPlanta centenas de vezes).
function bulkAddPlantas(p) {
  if (!p.itens || !p.itens.length) throw new Error("Nenhum item para importar.");
  const sheet = getPlantasSheet();
  const startRow = Math.max(sheet.getLastRow() + 1, PLANTA_DATA_START_ROW);
  const rows = p.itens.map(function (item) {
    return [item.nome || "", Number(item.preco) || 0, item.categoria || ""];
  });
  sheet.getRange(startRow, 1, rows.length, 3).setValues(rows);
  return { inseridos: rows.length };
}

// ---------------------------------------------------------------
// IMPORTAÇÃO EM MASSA DO CATÁLOGO DE PLANTAS (rodar 1 vez só)
// Como usar: abra este arquivo no editor do Apps Script, escolha a
// função "seedPlantasCatalogo" no menu suspenso ao lado de "Depurar/
// Executar" (no topo do editor) e clique em Executar. Ele pula
// qualquer planta que já exista no cadastro (não duplica se rodar
// de novo por engano).
// ---------------------------------------------------------------
function seedPlantasCatalogo() {
  const dados = [
    ["Umbigo Bahia", 28.0, "Laranjeiras"],
    ["Umbigo - Tomási", 28.0, "Laranjeiras"],
    ["Umb. Monte Parnaso", 28.0, "Laranjeiras"],
    ["Natal de Umbigo", 28.0, "Laranjeiras"],
    ["Natal - Suco", 28.0, "Laranjeiras"],
    ["Valência - Suco", 28.0, "Laranjeiras"],
    ["Laranja do Céu", 28.0, "Laranjeiras"],
    ["Laranja Comum", 28.0, "Laranjeiras"],
    ["Laranja Pêra", 28.0, "Laranjeiras"],
    ["Laranja da Índia", 28.0, "Laranjeiras"],
    ["Cara-cara", 28.0, "Laranjeiras"],
    ["Sanguínea", 28.0, "Laranjeiras"],
    ["Folha Murcha", 28.0, "Laranjeiras"],
    ["Champanhe", 28.0, "Laranjeiras"],
    ["Carambola", 28.0, "Laranjeiras"],
    ["Rubi", 28.0, "Laranjeiras"],
    ["Açúcar", 28.0, "Laranjeiras"],
    ["Navelate - Umbigo", 28.0, "Laranjeiras s/ Semente"],
    ["Navelina - Umbigo", 28.0, "Laranjeiras s/ Semente"],
    ["Lanelate - Umbigo", 28.0, "Laranjeiras s/ Semente"],
    ["Salustiana", 28.0, "Laranjeiras s/ Semente"],
    ["Valência Midnight", 28.0, "Laranjeiras s/ Semente"],
    ["Valência Delta Seedles", 28.0, "Laranjeiras s/ Semente"],
    ["Valência Late", 28.0, "Laranjeiras s/ Semente"],
    ["Ponkan", 28.0, "Bergamoteiras"],
    ["Danci", 28.0, "Bergamoteiras"],
    ["Do Céu", 28.0, "Bergamoteiras"],
    ["Caí Doce", 28.0, "Bergamoteiras"],
    ["Comum", 28.0, "Bergamoteiras"],
    ["Montenegrina", 28.0, "Bergamoteiras"],
    ["Montenegrina Rainha", 28.0, "Bergamoteiras"],
    ["Morgote", 28.0, "Bergamoteiras"],
    ["Decopon", 28.0, "Bergamoteiras"],
    ["Do Rio", 28.0, "Bergamoteiras"],
    ["Ortanique - Híbrida", 28.0, "Bergamoteiras s/ Semente"],
    ["Clemenules", 28.0, "Bergamoteiras s/ Semente"],
    ["Okitsú", 28.0, "Bergamoteiras s/ Semente"],
    ["Nova", 28.0, "Bergamoteiras s/ Semente"],
    ["Ceciliano", 28.0, "Limoeiros"],
    ["Taiti", 28.0, "Limoeiros"],
    ["Caviar", 28.0, "Limoeiros"],
    ["Cravo", 28.0, "Limoeiros"],
    ["Galego", 28.0, "Limoeiros"],
    ["Pércia", 28.0, "Limeiras"],
    ["Gigante", 28.0, "Limeiras"],
    ["Micado", 28.0, "Caquizeiros"],
    ["Kurat", 28.0, "Caquizeiros"],
    ["Fuyu", 28.0, "Caquizeiros"],
    ["Kiotto", 28.0, "Caquizeiros"],
    ["Eva", 28.0, "Macieiras"],
    ["Ana", 28.0, "Macieiras"],
    ["Galla", 28.0, "Macieiras"],
    ["Fugi", 28.0, "Macieiras"],
    ["Julieta", 28.0, "Macieiras"],
    ["Kaifer", 28.0, "Pereiras"],
    ["Leconte", 28.0, "Pereiras"],
    ["Schimit", 28.0, "Pereiras"],
    ["Hossuí", 28.0, "Pereiras"],
    ["Triunfo", 28.0, "Pereiras"],
    ["Kelsey", 28.0, "Ameixas"],
    ["Coração de Boi", 28.0, "Ameixas"],
    ["Bold Blum", 28.0, "Ameixas"],
    ["Santa Rosa", 28.0, "Ameixas"],
    ["Letícia Vermelha", 28.0, "Ameixas"],
    ["Pluma Sete", 28.0, "Ameixas"],
    ["Polirosa", 28.0, "Ameixas"],
    ["Amarelinha", 28.0, "Ameixas"],
    ["Vermelha", 28.0, "Ameixas"],
    ["Italianinha", 28.0, "Ameixas"],
    ["Irati", 28.0, "Ameixas"],
    ["Reubenel", 28.0, "Ameixas"],
    ["Fla 13-72", 28.0, "Pessegueiros"],
    ["Premier", 28.0, "Pessegueiros"],
    ["El Dourado", 28.0, "Pessegueiros"],
    ["Coral", 28.0, "Pessegueiros"],
    ["Nectarina", 28.0, "Pessegueiros"],
    ["Cardeal", 28.0, "Pessegueiros"],
    ["Diamante", 28.0, "Pessegueiros"],
    ["Cheripá", 28.0, "Pessegueiros"],
    ["Vila Nova", 28.0, "Pessegueiros"],
    ["Delicioso", 28.0, "Pessegueiros"],
    ["Precocinho", 28.0, "Pessegueiros"],
    ["Marli", 28.0, "Pessegueiros"],
    ["Chimarrita", 28.0, "Pessegueiros"],
    ["Cristal", 28.0, "Pessegueiros"],
    ["Granada", 28.0, "Pessegueiros"],
    ["Barbosa", 28.0, "Pessegueiros"],
    ["Fascínio", 28.0, "Pessegueiros"],
    ["Mandinho", 28.0, "Pessegueiros"],
    ["Rubimel", 28.0, "Pessegueiros"],
    ["Kampai", 28.0, "Pessegueiros"],
    ["Eragil", 28.0, "Pessegueiros"],
    ["Roxo de Valinhos", 28.0, "Figueiras"],
    ["Figo Branco", 28.0, "Figueiras"],
    ["Figo Amarelo", 28.0, "Figueiras"],
    ["Francês Preta", 28.0, "Videiras"],
    ["Niágara Branca", 28.0, "Videiras"],
    ["Niágara Rosa", 28.0, "Videiras"],
    ["Bordô", 28.0, "Videiras"],
    ["Isabel Preta Tardia", 28.0, "Videiras"],
    ["Isabel Precoce Preta", 28.0, "Videiras"],
    ["Rainha Itália", 28.0, "Videiras"],
    ["Isis", 28.0, "Videiras"],
    ["Vitória", 28.0, "Videiras"],
    ["Bananinha", 28.0, "Videiras"],
    ["Pitanga Paulista", 28.0, "Frutíferas Diversas"],
    ["Ariticum Gigante", 28.0, "Frutíferas Diversas"],
    ["Nogueira Pekan", 28.0, "Frutíferas Diversas"],
    ["Jaboticaba Comum", 28.0, "Frutíferas Diversas"],
    ["Guabijú", 28.0, "Frutíferas Diversas"],
    ["Cereja", 28.0, "Frutíferas Diversas"],
    ["Uvaia", 28.0, "Frutíferas Diversas"],
    ["Goiaba", 28.0, "Frutíferas Diversas"],
    ["Ameixa de Inverno", 28.0, "Frutíferas Diversas"],
    ["Marmelo do Japão", 28.0, "Frutíferas Diversas"],
    ["Abacate", 28.0, "Frutíferas Diversas"],
    ["Araçá", 28.0, "Frutíferas Diversas"],
    ["Romã", 28.0, "Frutíferas Diversas"],
    ["Guabiroba", 28.0, "Frutíferas Diversas"],
    ["Amora preta", 28.0, "Frutíferas Diversas"],
    ["Aroeira Periquita", 44.0, "Ornamentais e Sombra"],
    ["Acácia Régia", 44.0, "Ornamentais e Sombra"],
    ["Acácia Mimosa", 44.0, "Ornamentais e Sombra"],
    ["Canela", 44.0, "Ornamentais e Sombra"],
    ["Cereja Flor", 44.0, "Ornamentais e Sombra"],
    ["Espatódia", 44.0, "Ornamentais e Sombra"],
    ["Sibipiruna", 44.0, "Ornamentais e Sombra"],
    ["Espirradeira", 44.0, "Ornamentais e Sombra"],
    ["Jacarandá Azul", 44.0, "Ornamentais e Sombra"],
    ["Loro Tempero", 44.0, "Ornamentais e Sombra"],
    ["Ingá", 44.0, "Ornamentais e Sombra"],
    ["Flamboyant", 44.0, "Ornamentais e Sombra"],
    ["Pata de Vaca", 44.0, "Ornamentais e Sombra"],
    ["Tipuana", 44.0, "Ornamentais e Sombra"],
    ["Cobrina", 44.0, "Ornamentais e Sombra"],
    ["Ipê", 44.0, "Ornamentais e Sombra"],
    ["Chorão Salso", 44.0, "Ornamentais e Sombra"],
    ["Chuva de Ouro", 44.0, "Ornamentais e Sombra"],
    ["Plátano", 44.0, "Ornamentais e Sombra"],
    ["Palmeira Real", 44.0, "Ornamentais e Sombra"],
    ["Palma Leque", 44.0, "Ornamentais e Sombra"],
    ["Cinamomo - Sombra", 44.0, "Ornamentais e Sombra"],
    ["Jamelão - Sombra", 44.0, "Ornamentais e Sombra"],
    ["Hortência", 44.0, "Flores"],
    ["Dracena - Folhagem", 44.0, "Flores"],
    ["Azaléia Branca", 44.0, "Flores"],
    ["Azaléia", 44.0, "Flores"],
    ["Camélia", 44.0, "Flores"],
    ["Camélia Matizada", 44.0, "Flores"],
    ["Estremosa", 44.0, "Flores"],
    ["Chorão de flor", 44.0, "Flores"],
    ["Jasmim", 44.0, "Flores"],
    ["Buquê de Noiva", 44.0, "Flores"],
    ["Caleandra", 44.0, "Flores"],
    ["Hibisco", 44.0, "Flores"],
    ["Goiaba Tailandesa", 53.0, "Mudas Nobres"],
    ["Moreia Branca", 53.0, "Mudas Nobres"],
    ["Dracena Tricolor", 53.0, "Mudas Nobres"],
    ["Dracena Pleomele", 53.0, "Mudas Nobres"],
    ["Kiwi - Macho e Fêmea", 53.0, "Mudas Nobres"],
    ["Cipreste tuya (+/- 70 cm)", 53.0, "Mudas Nobres"],
    ["Palmeira Família", 53.0, "Mudas Nobres"],
    ["Fruta do Conde", 53.0, "Mudas Nobres"],
    ["Graviola", 53.0, "Mudas Nobres"],
    ["Canela Doce (+/- 1,50m)", 53.0, "Mudas Nobres"],
    ["Café - Fruto", 53.0, "Mudas Nobres"],
    ["Lágrima de Cristo", 53.0, "Mudas Nobres"],
    ["Quaresmeira", 53.0, "Mudas Nobres"],
    ["Pata de Elefante", 53.0, "Mudas Nobres"],
    ["Maracujá", 53.0, "Mudas Nobres"],
    ["Jaca", 53.0, "Mudas Nobres"],
    ["Fênix", 53.0, "Mudas Nobres"],
    ["Ingá (+/- 1,20m)", 53.0, "Mudas Nobres"],
    ["Acerola (Vitamina C)", 53.0, "Mudas Nobres"],
    ["Manga Enxertada", 53.0, "Mudas Nobres"],
    ["Jabuticaba Híbrida", 53.0, "Mudas Nobres"],
    ["Abacate Enxertado", 53.0, "Mudas Nobres"],
    ["Butiá Doce", 53.0, "Mudas Nobres"],
    ["Taquarinha de Jardim", 53.0, "Mudas Nobres"],
    ["Pinheirinho de Jardim", 53.0, "Mudas Nobres"],
    ["Podocarpus", 53.0, "Mudas Nobres"],
    ["Mirta - Chinesa Dupla", 53.0, "Mudas Nobres"],
    ["Hortência Japonesa", 53.0, "Mudas Nobres"],
    ["Ficus", 53.0, "Mudas Nobres"],
    ["Manacá da Serra", 53.0, "Mudas Nobres"],
    ["Carambola (Muda Nobre)", 53.0, "Mudas Nobres"],
    ["Buchinho", 53.0, "Mudas Nobres"],
    ["Cróton", 53.0, "Mudas Nobres"],
    ["Estrelitsa", 53.0, "Mudas Nobres"],
    ["Alamanda", 53.0, "Mudas Nobres"],
    ["Ráfia", 53.0, "Mudas Nobres"],
    ["Oliveira", 53.0, "Mudas Nobres"],
    ["Três Maria", 53.0, "Mudas Nobres"],
    ["Jabuticaba Gr. Lata (+/-1m)", 150.0, "Lançamentos"],
    ["Figueira Chilena - Sombra", 35.0, "Lançamentos"],
    ["Mertilo", 35.0, "Lançamentos"],
    ["Nogueira Pekan Enxertada", 105.0, "Lançamentos"],
    ["Moranguinho de Árvore", 35.0, "Lançamentos"],
    ["Tuya Europa", 35.0, "Lançamentos"],
    ["Palmeira Real (+/- 2 m)", 105.0, "Lançamentos"],
    ["Cipreste vela", 105.0, "Lançamentos"],
    ["Pau Brasil", 35.0, "Lançamentos"],
    ["Pitaya", 35.0, "Lançamentos"],
    ["Coqueiro Jerivá", 105.0, "Lançamentos"],
    ["Kaizuka", 150.0, "Lançamentos"],
    ["Palmeira Areca", 105.0, "Lançamentos"],
    ["Acer", 35.0, "Lançamentos"],
    ["Damasco", 35.0, "Lançamentos"],
    ["Atemóia", 35.0, "Lançamentos"],
    ["Nós Macadamia", 35.0, "Lançamentos"],
    ["Magnólia Amarela", 35.0, "Lançamentos"],
    ["Magnólia Roxa", 35.0, "Lançamentos"],
    ["Hibisco Híbrido", 35.0, "Lançamentos"],
    ["Rosa do Deserto", 105.0, "Lançamentos"],
    ["Jasmim Manga", 105.0, "Lançamentos"],
    ["Abacate Avocado", 35.0, "Lançamentos"],
    ["Castanha Portuguesa", 35.0, "Lançamentos"],
    ["Cravo da Índia", 35.0, "Lançamentos"],
    ["Banana", 35.0, "Lançamentos"],
    ["Glicínia Azul", 35.0, "Lançamentos"],
    ["Glicínia Rosa", 35.0, "Lançamentos"],
    ["01 - Amarelo Vivo", 28.0, "Roseiras"],
    ["02 - Vermelho Veludo Brilhante", 28.0, "Roseiras"],
    ["03 - Cor de Rosa", 28.0, "Roseiras"],
    ["04 - Telha", 28.0, "Roseiras"],
    ["07 - Vermelho Carmim", 28.0, "Roseiras"],
    ["09 - Vermelho Fogo", 28.0, "Roseiras"],
    ["12 - Matizada", 28.0, "Roseiras"],
    ["13 - Vermelha e Branca", 28.0, "Roseiras"],
    ["14 - Branca", 28.0, "Roseiras"],
    ["15 - Granada três cores", 28.0, "Roseiras"],
    ["16 - Azul", 28.0, "Roseiras"],
    ["17 - Cenoura", 28.0, "Roseiras"],
    ["18 - Diorana Laranja", 28.0, "Roseiras"],
    ["Erva Mate - uni", 10.0, "Árvores p/ Reflorestamento"],
    ["Eucaliptus Grandis", 177.0, "Árvores p/ Reflorestamento"],
    ["Eucaliptus Dunnii", 177.0, "Árvores p/ Reflorestamento"],
    ["Eucaliptus Clonado", 354.0, "Árvores p/ Reflorestamento"],
    ["Eucaliptus Saligno", 177.0, "Árvores p/ Reflorestamento"],
    ["Pinus", 177.0, "Árvores p/ Reflorestamento"],
    ["Acácia Negra", 177.0, "Árvores p/ Reflorestamento"],
    ["Cipreste de Cerca", 213.0, "Árvores p/ Reflorestamento"],
    ["Pingo de Ouro", 35.0, "Árvores p/ Reflorestamento"],
    ["Mini Azaléia", 130.0, "Árvores p/ Reflorestamento"],
    ["Mini Rosa", 130.0, "Árvores p/ Reflorestamento"],
    ["Flores (Reflorestamento)", 35.0, "Árvores p/ Reflorestamento"],
    ["Mudas Nativas - cx. 50uni", 250.0, "Árvores p/ Reflorestamento"],
    ["Saco de Pedras", 35.0, "Árvores p/ Reflorestamento"],
    ["Substrato", 35.0, "Árvores p/ Reflorestamento"],
    ["Casca de Pinus", 35.0, "Árvores p/ Reflorestamento"],
    ["Moranguinho Maço com 25 unidades", 35.0, "Árvores p/ Reflorestamento"]
  ];

  const sheet = getPlantasSheet();
  const existentes = listPlantas().map(function (p) { return p.nome.trim().toLowerCase(); });

  const novas = dados.filter(function (d) {
    return existentes.indexOf(String(d[0]).trim().toLowerCase()) === -1;
  });

  if (novas.length === 0) {
    Logger.log("Nenhuma planta nova para importar (todas já existem no cadastro).");
    return;
  }

  const startRow = Math.max(sheet.getLastRow() + 1, PLANTA_DATA_START_ROW);
  const rows = novas.map(function (d) { return [d[0], d[1], d[2]]; });
  sheet.getRange(startRow, 1, rows.length, 3).setValues(rows);

  Logger.log(novas.length + " plantas importadas com sucesso.");
}

function getPedidosPlantasSheet() {
  let sheet = ss().getSheetByName(PEDIDOS_PLANTAS_SHEET);
  if (!sheet) {
    sheet = ss().insertSheet(PEDIDOS_PLANTAS_SHEET);
    const headers = PLANTA_PEDIDO_COLS.concat(["Status", "Data Criacao", "Pedido ID"]);
    sheet.getRange(PLANTA_HEADER_ROW, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

function listPedidosPlantas() {
  const sheet = getPedidosPlantasSheet();
  const totalCols = PLANTA_PEDIDO_COLS.length + 3;
  const maxRow = sheet.getLastRow();
  if (maxRow < PLANTA_DATA_START_ROW) return [];
  const values = sheet.getRange(PLANTA_DATA_START_ROW, 1, maxRow - PLANTA_DATA_START_ROW + 1, totalCols).getValues();
  const all = [];
  values.forEach(function (row, i) {
    if (!row[0] || isPlaceholderRow(row[0])) return;
    const rowIndex = PLANTA_DATA_START_ROW + i;
    const pedidoIdRaw = row[9];
    all.push({
      rowIndex: rowIndex,
      nome: row[0],
      telefone: normalizeStr(row[1]),
      endereco: normalizeStr(row[2]),
      municipio: normalizeStr(row[3]),
      planta: normalizeStr(row[4]),
      qtd: row[5] || 0,
      observacoes: normalizeStr(row[6]),
      status: row[7] || "Ativo",
      dataCriacao: row[8] ? formatDate(row[8]) : "",
      pedidoId: pedidoIdRaw ? String(pedidoIdRaw) : ("planta#" + rowIndex)
    });
  });
  return all;
}

function buildPedidoPlantaRow(p) {
  return [p.nome || "", p.telefone || "", p.endereco || "", p.municipio || "", p.planta || "", p.qtd || 0, p.observacoes || ""];
}

function addPedidoPlantaMulti(p) {
  if (!p.itens || !p.itens.length) throw new Error("Informe ao menos uma planta no pedido.");
  const pedidoId = p.pedidoId || Utilities.getUuid();
  const agora = new Date();
  const sheet = getPedidosPlantasSheet();
  const criados = [];
  p.itens.forEach(function (item) {
    const newRow = Math.max(sheet.getLastRow() + 1, PLANTA_DATA_START_ROW);
    const rowData = buildPedidoPlantaRow({
      nome: p.nome, telefone: p.telefone, endereco: p.endereco, municipio: p.municipio,
      planta: item.planta, qtd: item.qtd, observacoes: p.observacoes
    });
    sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
    sheet.getRange(newRow, rowData.length + 1, 1, 3).setValues([["Ativo", agora, pedidoId]]);
    criados.push({ rowIndex: newRow });
  });
  return { pedidoId: pedidoId, itens: criados };
}

function updatePedidoPlantaMulti(p) {
  if (!p.pedidoId) throw new Error("Pedido inválido.");
  if (!p.itens || !p.itens.length) throw new Error("Informe ao menos uma planta no pedido.");

  const existentes = findPedidoPlantaGroupRows(p.pedidoId);
  const dataOriginal = existentes.length && existentes[0].dataCriacaoRaw ? existentes[0].dataCriacaoRaw : new Date();
  const statusOriginal = existentes.length ? existentes[0].status : "Ativo";

  deletePedidoPlantaGroup({ pedidoId: p.pedidoId });

  const sheet = getPedidosPlantasSheet();
  const criados = [];
  p.itens.forEach(function (item) {
    const newRow = Math.max(sheet.getLastRow() + 1, PLANTA_DATA_START_ROW);
    const rowData = buildPedidoPlantaRow({
      nome: p.nome, telefone: p.telefone, endereco: p.endereco, municipio: p.municipio,
      planta: item.planta, qtd: item.qtd, observacoes: p.observacoes
    });
    sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
    sheet.getRange(newRow, rowData.length + 1, 1, 3).setValues([[statusOriginal, dataOriginal, p.pedidoId]]);
    criados.push({ rowIndex: newRow });
  });
  return { pedidoId: p.pedidoId, itens: criados };
}

function findPedidoPlantaGroupRows(pedidoId) {
  const sheet = getPedidosPlantasSheet();
  const idCol = PLANTA_PEDIDO_COLS.length + 3;
  const statusCol = PLANTA_PEDIDO_COLS.length + 1;
  const dataCol = PLANTA_PEDIDO_COLS.length + 2;
  const maxRow = sheet.getLastRow();
  const out = [];
  if (maxRow < PLANTA_DATA_START_ROW) return out;
  const ids = sheet.getRange(PLANTA_DATA_START_ROW, idCol, maxRow - PLANTA_DATA_START_ROW + 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    const rowIndex = PLANTA_DATA_START_ROW + i;
    const rowId = ids[i][0] ? String(ids[i][0]) : ("planta#" + rowIndex);
    if (rowId === pedidoId) {
      const statusVal = sheet.getRange(rowIndex, statusCol).getValue() || "Ativo";
      const dataVal = sheet.getRange(rowIndex, dataCol).getValue();
      out.push({ rowIndex: rowIndex, status: statusVal, dataCriacaoRaw: dataVal });
    }
  }
  return out;
}

function deletePedidoPlantaGroup(p) {
  if (!p.pedidoId) throw new Error("Pedido inválido.");
  const sheet = getPedidosPlantasSheet();
  const rows = findPedidoPlantaGroupRows(p.pedidoId);
  const indices = rows.map(function (r) { return r.rowIndex; }).sort(function (a, b) { return b - a; });
  indices.forEach(function (rowIndex) { sheet.deleteRow(rowIndex); });
  return { deleted: true, count: rows.length };
}

function setPedidoPlantaGroupStatus(p, status) {
  if (!p.pedidoId) throw new Error("Pedido inválido.");
  const sheet = getPedidosPlantasSheet();
  const statusCol = PLANTA_PEDIDO_COLS.length + 1;
  const rows = findPedidoPlantaGroupRows(p.pedidoId);
  rows.forEach(function (r) { sheet.getRange(r.rowIndex, statusCol).setValue(status); });
  return { status: status, count: rows.length };
}

// ---------------------------------------------------------------
// MÁQUINAS — cadastro (nome/marca/valor por dia) e aluguéis
// ---------------------------------------------------------------
const MAQUINA_HEADER_ROW = 1;
const MAQUINA_DATA_START_ROW = 2;
const MAQUINAS_SHEET = "Maquinas";
const ALUGUEIS_SHEET = "Alugueis";
const ALUGUEL_COLS = ["Maquina", "Cliente", "Telefone", "Endereco", "Municipio", "DataInicio", "DataFim"];

function getMaquinasSheet() {
  let sheet = ss().getSheetByName(MAQUINAS_SHEET);
  if (!sheet) {
    sheet = ss().insertSheet(MAQUINAS_SHEET);
    sheet.getRange(MAQUINA_HEADER_ROW, 1, 1, 3).setValues([["Nome", "Marca", "ValorDia"]]);
  }
  return sheet;
}

function listMaquinas() {
  const sheet = getMaquinasSheet();
  const maxRow = sheet.getLastRow();
  if (maxRow < MAQUINA_DATA_START_ROW) return [];
  const values = sheet.getRange(MAQUINA_DATA_START_ROW, 1, maxRow - MAQUINA_DATA_START_ROW + 1, 3).getValues();
  const out = [];
  values.forEach(function (row, i) {
    if (!row[0] || isPlaceholderRow(row[0])) return;
    out.push({ rowIndex: MAQUINA_DATA_START_ROW + i, nome: row[0], marca: normalizeStr(row[1]), valorDia: Number(row[2]) || 0 });
  });
  return out;
}

function addMaquina(p) {
  const sheet = getMaquinasSheet();
  const newRow = Math.max(sheet.getLastRow() + 1, MAQUINA_DATA_START_ROW);
  sheet.getRange(newRow, 1, 1, 3).setValues([[p.nome || "", p.marca || "", Number(p.valorDia) || 0]]);
  return { rowIndex: newRow };
}

function updateMaquina(p) {
  const sheet = getMaquinasSheet();
  if (!p.rowIndex || p.rowIndex < MAQUINA_DATA_START_ROW) throw new Error("Linha inválida.");
  sheet.getRange(p.rowIndex, 1, 1, 3).setValues([[p.nome || "", p.marca || "", Number(p.valorDia) || 0]]);
  return { rowIndex: p.rowIndex };
}

function deleteMaquina(p) {
  const sheet = getMaquinasSheet();
  if (!p.rowIndex || p.rowIndex < MAQUINA_DATA_START_ROW) throw new Error("Linha inválida.");
  sheet.deleteRow(p.rowIndex);
  return { deleted: true };
}

function getAlugueisSheet() {
  let sheet = ss().getSheetByName(ALUGUEIS_SHEET);
  if (!sheet) {
    sheet = ss().insertSheet(ALUGUEIS_SHEET);
    const headers = ALUGUEL_COLS.concat(["Status", "DataCriacao"]);
    sheet.getRange(MAQUINA_HEADER_ROW, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

function listAlugueis() {
  const sheet = getAlugueisSheet();
  const totalCols = ALUGUEL_COLS.length + 2;
  const maxRow = sheet.getLastRow();
  if (maxRow < MAQUINA_DATA_START_ROW) return [];
  const values = sheet.getRange(MAQUINA_DATA_START_ROW, 1, maxRow - MAQUINA_DATA_START_ROW + 1, totalCols).getValues();
  const out = [];
  values.forEach(function (row, i) {
    if (!row[0] || isPlaceholderRow(row[0])) return;
    const rowIndex = MAQUINA_DATA_START_ROW + i;
    out.push({
      rowIndex: rowIndex,
      maquina: row[0],
      cliente: row[1],
      telefone: normalizeStr(row[2]),
      endereco: normalizeStr(row[3]),
      municipio: normalizeStr(row[4]),
      dataInicio: row[5] ? formatDateSimples(row[5]) : "",
      dataFim: row[6] ? formatDateSimples(row[6]) : "",
      status: row[7] || "Ativo",
      dataCriacao: row[8] ? formatDate(row[8]) : ""
    });
  });
  return out;
}

function buildAluguelRow(p) {
  return [
    p.maquina || "", p.cliente || "", p.telefone || "", p.endereco || "", p.municipio || "",
    p.dataInicio ? new Date(p.dataInicio + "T00:00:00") : "",
    p.dataFim ? new Date(p.dataFim + "T00:00:00") : ""
  ];
}

function addAluguel(p) {
  if (!p.maquina) throw new Error("Escolha uma máquina.");
  if (!p.dataInicio || !p.dataFim) throw new Error("Informe as datas de início e devolução.");
  const sheet = getAlugueisSheet();
  const newRow = Math.max(sheet.getLastRow() + 1, MAQUINA_DATA_START_ROW);
  const rowData = buildAluguelRow(p);
  sheet.getRange(newRow, 1, 1, rowData.length).setValues([rowData]);
  sheet.getRange(newRow, rowData.length + 1, 1, 2).setValues([["Ativo", new Date()]]);
  return { rowIndex: newRow };
}

function updateAluguel(p) {
  if (!p.rowIndex || p.rowIndex < MAQUINA_DATA_START_ROW) throw new Error("Linha inválida.");
  const sheet = getAlugueisSheet();
  const rowData = buildAluguelRow(p);
  sheet.getRange(p.rowIndex, 1, 1, rowData.length).setValues([rowData]);
  return { rowIndex: p.rowIndex };
}

function deleteAluguel(p) {
  if (!p.rowIndex || p.rowIndex < MAQUINA_DATA_START_ROW) throw new Error("Linha inválida.");
  getAlugueisSheet().deleteRow(p.rowIndex);
  return { deleted: true };
}

function setAluguelStatus(p, status) {
  if (!p.rowIndex || p.rowIndex < MAQUINA_DATA_START_ROW) throw new Error("Linha inválida.");
  const statusCol = ALUGUEL_COLS.length + 1;
  getAlugueisSheet().getRange(p.rowIndex, statusCol).setValue(status);
  return { rowIndex: p.rowIndex, status: status };
}

// ---------------------------------------------------------------
// BACKUP AUTOMÁTICO (copia a planilha inteira pra uma pasta separada
// no Drive, todo dia, mantendo só os 7 backups mais recentes)
// ---------------------------------------------------------------
const BACKUP_PASTA_NOME = "Backups Aves Agro";
const BACKUP_MAXIMO = 7;

// ---------------------------------------------------------------
// AJUDA PRA TESTAR EM CÓPIA (não faz parte do app -- só rode manualmente
// aqui no editor quando quiser levar a lista de tipos de ave personalizados
// pra uma planilha de teste)
// ---------------------------------------------------------------

// Rode ISSO na planilha REAL (produção). Depois, olhe em Ver → Registros de
// execução (ou Execuções, no menu lateral) e copie o texto que aparecer.
function exportarRegistroCategorias() {
  const texto = JSON.stringify(getCategoriasAvesRegistro());
  Logger.log(texto);
  return texto;
}

// Rode ISSO na planilha de TESTE. Cole o texto copiado do passo acima
// dentro das aspas abaixo, substituindo o exemplo, antes de rodar.
function importarRegistroCategorias() {
  const textoColado = 'COLE_AQUI_O_TEXTO_COPIADO';
  if (textoColado === 'COLE_AQUI_O_TEXTO_COPIADO') {
    Logger.log("Cole o texto copiado de exportarRegistroCategorias() na variável textoColado antes de rodar.");
    return;
  }
  const lista = JSON.parse(textoColado);
  salvarCategoriasAvesRegistro(lista);
  Logger.log("Importado! " + lista.length + " tipos de ave configurados nessa planilha de teste.");
}

// Rode essa função UMA VEZ manualmente (Selecionar função → Executar,
// aqui no próprio editor do Apps Script) pra ligar o backup automático.
// Depois disso ela roda sozinha todo dia, sem precisar fazer mais nada.
function configurarBackupAutomatico() {
  const gatilhos = ScriptApp.getProjectTriggers();
  gatilhos.forEach(function (g) {
    if (g.getHandlerFunction() === "fazerBackupDiario") {
      ScriptApp.deleteTrigger(g);
    }
  });

  ScriptApp.newTrigger("fazerBackupDiario")
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();

  Logger.log("Backup automático configurado! Vai rodar todo dia por volta das 3h da manhã, guardando sempre os últimos " + BACKUP_MAXIMO + " backups (~1 semana).");
}

function fazerBackupDiario() {
  const origem = SpreadsheetApp.getActiveSpreadsheet();
  const nomeBackup = "Backup - " + origem.getName() + " - " +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT-3", "dd-MM-yyyy HH'h'mm");

  const pasta = getOuCriarPastaBackup();
  const arquivoOrigem = DriveApp.getFileById(origem.getId());
  arquivoOrigem.makeCopy(nomeBackup, pasta);

  limparBackupsAntigos(pasta, BACKUP_MAXIMO);
}

function getOuCriarPastaBackup() {
  const pastas = DriveApp.getFoldersByName(BACKUP_PASTA_NOME);
  if (pastas.hasNext()) return pastas.next();
  return DriveApp.createFolder(BACKUP_PASTA_NOME);
}

// Mantém só os "maximo" backups mais recentes, apagando o resto,
// independente de quantos dias eles têm.
function limparBackupsAntigos(pasta, maximo) {
  const arquivos = [];
  const iterador = pasta.getFilesByType(MimeType.GOOGLE_SHEETS);
  while (iterador.hasNext()) {
    arquivos.push(iterador.next());
  }
  arquivos.sort(function (a, b) {
    return b.getDateCreated() - a.getDateCreated(); // mais novo primeiro
  });
  for (let i = maximo; i < arquivos.length; i++) {
    arquivos[i].setTrashed(true);
  }
}

// ---------------------------------------------------------------
// UTIL
// ---------------------------------------------------------------
function requireRow(rowIndex) {
  if (!rowIndex || rowIndex < DATA_START_ROW) throw new Error("Linha inválida.");
}

function normalizeStr(v) {
  if (v === null || v === undefined) return "";
  return String(v);
}

// Detecta linhas de exemplo/modelo (onde a célula do nome é literalmente
// "Nome", "Cliente", etc.) para não aparecerem como cadastros de verdade.
function isPlaceholderRow(nomeValue) {
  const v = String(nomeValue).trim().toLowerCase();
  return v === "nome" || v === "cliente" || v === "nome do cliente" || v === "exemplo";
}

function formatDate(v) {
  try {
    return Utilities.formatDate(new Date(v), Session.getScriptTimeZone() || "GMT-3", "dd/MM/yyyy HH:mm");
  } catch (e) {
    return "";
  }
}

function formatDateSimples(v) {
  try {
    return Utilities.formatDate(new Date(v), Session.getScriptTimeZone() || "GMT-3", "dd/MM/yyyy");
  } catch (e) {
    return "";
  }
}

// Acha a última linha (a partir de startRow) que tem valor na coluna A,
// ignorando formatação/linhas em branco no meio ou no fim da aba.
function lastNonEmptyRow(sheet, col, startRow) {
  const lastPossible = sheet.getMaxRows();
  if (lastPossible < startRow) return startRow - 1;
  const values = sheet.getRange(startRow, col, lastPossible - startRow + 1, 1).getValues();
  let last = startRow - 1;
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] !== "" && values[i][0] !== null) last = startRow + i;
  }
  return last;
}