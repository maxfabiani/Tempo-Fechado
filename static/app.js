
// v8.21.38 - Administracao: exclusao protegida de usuarios cadastrados.
function marcarGuiaAtivaV81738(nomeGuia) {
  try {
    const botoes = document.querySelectorAll(".menu button, aside button, nav button");
    botoes.forEach(btn => {
      btn.classList.remove("active", "ativo", "selected", "selecionado");
      btn.removeAttribute("aria-current");
    });
    const alvoTexto = nomeGuia === "manual" ? "processamento manual" : "processamento";
    botoes.forEach(btn => {
      const texto = (btn.textContent || "").trim().toLowerCase();
      if (texto === alvoTexto) {
        btn.classList.add("active");
        btn.setAttribute("aria-current", "page");
      }
    });
  } catch (e) {}
}
function limparEstadoOperacionalV81738() {
  try {
    const logManual = document.getElementById("manualRoboLogPanel");
    if (logManual) logManual.remove();
  } catch (e) {}
}


const qs = (id) => document.getElementById(id);

function limparSubtitulosGuiasV82130() {
  const alvo = qs("subtituloPagina");
  if (!alvo) return;
  if (alvo.textContent || alvo.innerHTML) {
    alvo.textContent = "";
    alvo.innerHTML = "";
  }
  alvo.setAttribute("data-subtitulos-removidos", "v8.21.38");
}

document.addEventListener("DOMContentLoaded", () => {
  limparSubtitulosGuiasV82130();
  const alvo = qs("subtituloPagina");
  if (!alvo || !window.MutationObserver) return;
  const observer = new MutationObserver(() => limparSubtitulosGuiasV82130());
  observer.observe(alvo, { childList: true, characterData: true, subtree: true });
});

let dadosConsolidadoAtual = [];
let ordenacaoSaldo = null; // null | "asc" | "desc"
let ordenacaoHE = null;    // null | "asc" | "desc"
let ordenacaoInconsistencias = null; // null | "asc" | "desc"
let paginaAtual = "consolidado";
let datasFiltro = [];
let datasDashboardFiltro = [];

const LIMITE_RENDER_TABELA = 800;
const TAMANHO_LOTE_RENDER_TABELA = 120;
const TIMEOUT_CONSULTA_ESPELHO_MS = 45000; // v8.11.25: primeira leitura/normalizacao do Excel pode passar de 12s em bases grandes
let controllersNavegacaoAtivos = [];
let tabelaRenderToken = 0;
const cacheOpcoesFiltros = new Map();
const PAGINAS_PAINEL_COMPARTILHADO_V81744 = new Set([
  "dashboard_executivo",
  "alertas_automaticos",
  "auditoria_premium",
  "score_operacional",
  "processamento_integrado",
  "processamento_manual_robo",
  "central_operacao",
  "diagnostico_instalacao",
  "painel_implantacao",
  "administracao",
  "configuracao_robo",
  "admin_config_robo"
]);

const TITULOS_PAINEL_COMPARTILHADO_V81744 = {
  dashboard_executivo: "Dashboard Executivo",
  alertas_automaticos: "Central de Alertas",
  auditoria_premium: "Auditoria",
  score_operacional: "Score Operacional",
  processamento_integrado: "Processamento",
  processamento_manual_robo: "Processamento Manual",
  central_operacao: "Central de Operação",
  diagnostico_instalacao: "Diagnóstico",
  painel_implantacao: "Painel de Implantação",
  administracao: "Administração",
  configuracao_robo: "Configuração do Robô",
  admin_config_robo: "Administração do Robô"
};




// v8.21.4 - Blindagem definitiva do subtitulo da Central de Alertas.
// Algumas instalacoes mantinham textos antigos por cache ou por renderizacao tardia.
// Esta rotina limpa o subtitulo da topbar e remove qualquer paragrafo legado da Central.
function limparSubtituloCentralAlertasV8214() {
  try {
    const alvo = qs("subtituloPagina");
    if (paginaAtual === "alertas_automaticos" && alvo) {
      alvo.textContent = "";
      alvo.innerHTML = "";
      alvo.setAttribute("data-central-alertas-limpo", "v8.21.9");
    }
    if (paginaAtual !== "alertas_automaticos") return;
    const frasesLegadas = [
      "fase 02",
      "riscos priorizados",
      "acao gerencial",
      "ação gerencial",
      "sobre banco, h.e.",
      "marcações e jornada",
      "marcacoes e jornada"
    ];
    document.querySelectorAll("#subtituloPagina, .exec-panel-title p, .dashboard-hero p, .alertas-panel p, p, small").forEach(el => {
      if (!el || !el.textContent) return;
      const texto = el.textContent.toLowerCase();
      const achou = frasesLegadas.some(f => texto.includes(f));
      if (achou) {
        el.textContent = "";
        el.innerHTML = "";
        el.classList.add("hidden");
        el.setAttribute("data-removido-v8214", "subtitulo-central-alertas");
      }
    });
  } catch (e) {}
}

function ativarBlindagemSubtituloCentralAlertasV8214() {
  try {
    const alvo = qs("subtituloPagina");
    if (!alvo || alvo.__blindagemCentralAlertasV8214) return;
    alvo.__blindagemCentralAlertasV8214 = true;
    const obs = new MutationObserver(() => {
      if (paginaAtual === "alertas_automaticos") limparSubtituloCentralAlertasV8214();
    });
    obs.observe(alvo, { childList: true, characterData: true, subtree: true });
  } catch (e) {}
}

function limparCacheOpcoesFiltros() {
  cacheOpcoesFiltros.clear();
}

async function obterOpcoesFiltros(cc = "", force = false) {
  const chave = String(cc || "");
  if (!force && cacheOpcoesFiltros.has(chave)) {
    return cacheOpcoesFiltros.get(chave);
  }

  const p = new URLSearchParams();
  if (cc) p.set("cc", cc);
  if (force) {
    p.set("refresh", "1");
    p.set("_", String(Date.now()));
  }
  const url = `/api/opcoes-filtros${p.toString() ? "?" + p.toString() : ""}`;
  const resp = await fetch(url);

  if (resp.status === 401) {
    window.location.href = "/login";
    return { ccs: [], nomes: [], turnos: [] };
  }

  const json = await resp.json();
  const opcoes = {
    ccs: json.ccs || [],
    nomes: json.nomes || [],
    turnos: json.turnos || []
  };
  cacheOpcoesFiltros.set(chave, opcoes);
  return opcoes;
}

function limitarRenderTabela(rows, contexto = "registros") {
  const lista = Array.isArray(rows) ? rows : [];
  if (lista.length <= LIMITE_RENDER_TABELA) return lista;
  const status = qs("statusCarga");
  if (status) {
    status.textContent = `${lista.length} ${contexto} encontrados. Exibindo os primeiros ${LIMITE_RENDER_TABELA} para preservar a navegação. Use filtros para refinar.`;
  }
  return lista.slice(0, LIMITE_RENDER_TABELA);
}

function criarControllerNavegacao(timeoutMs = 12000) {
  const controller = new AbortController();
  controllersNavegacaoAtivos.push(controller);
  const timer = setTimeout(() => {
    try { controller.abort(); } catch (e) {}
  }, timeoutMs);
  return { controller, timer };
}

function finalizarControllerNavegacao(controller, timer) {
  clearTimeout(timer);
  controllersNavegacaoAtivos = controllersNavegacaoAtivos.filter(c => c !== controller);
}

function abortarRequisicoesNavegacao() {
  controllersNavegacaoAtivos.forEach(c => { try { c.abort(); } catch (e) {} });
  controllersNavegacaoAtivos = [];
}

function cancelarRenderizacaoTabela() {
  tabelaRenderToken += 1;
}

function aguardarProximoFrame() {
  return new Promise(resolve => {
    const raf = window.requestAnimationFrame || ((cb) => setTimeout(cb, 16));
    raf(resolve);
  });
}

async function renderTabelaEmLotes(tbodyId, rows, renderLinha, opcoes = {}) {
  const tbody = qs(tbodyId);
  if (!tbody) return;

  const token = ++tabelaRenderToken;
  const lista = Array.isArray(rows) ? rows : [];
  const lote = Number(opcoes.lote || TAMANHO_LOTE_RENDER_TABELA);

  tbody.innerHTML = "";
  if (!lista.length) {
    if (opcoes.htmlVazio) tbody.innerHTML = opcoes.htmlVazio;
    if (typeof opcoes.aoConcluir === "function") opcoes.aoConcluir(0);
    return;
  }

  for (let i = 0; i < lista.length; i += lote) {
    if (token !== tabelaRenderToken) return;
    const html = lista.slice(i, i + lote).map(renderLinha).join("");
    tbody.insertAdjacentHTML("beforeend", html);
    if (i + lote < lista.length) await aguardarProximoFrame();
  }

  if (token === tabelaRenderToken && typeof opcoes.aoConcluir === "function") {
    opcoes.aoConcluir(lista.length);
  }
}

// v8.10.45 - controle de requisições longas por guia
let controllerExtratoBancoHoras = null;
let timerFiltroPrincipal = null;


// v8.10.29 - Estabilização de navegação entre guias
// Cada troca de guia recebe um token. Respostas assíncronas antigas são ignoradas
// para impedir que o conteúdo de uma guia apareça em outra quando uma consulta demora.
let navegacaoToken = 0;

function iniciarTransicaoPagina(pagina) {
  abortarRequisicoesNavegacao();
  navegacaoToken += 1;
  limparAreaConteudoTransicao(pagina);
  return navegacaoToken;
}

function tokenAtualPagina() {
  return navegacaoToken;
}

function requisicaoAindaValida(paginaEsperada, tokenEsperado) {
  return paginaAtual === paginaEsperada && navegacaoToken === tokenEsperado;
}

function limparAreaConteudoTransicao(pagina) {
  cancelarRenderizacaoTabela();
  const status = qs("statusCarga");
  const tbody = qs("tbody");
  const tbodyDashboard = qs("tbodyDashboard");
  const painelPremium = qs("dashboardExecutivoPremium");
  const diagnostico = qs("diagnosticoBancoHoras");

  if (status) {
    status.classList.remove("bh2-header-controls");
    status.textContent = "Carregando...";
  }
  if (tbody) tbody.innerHTML = "";
  if (tbodyDashboard) tbodyDashboard.innerHTML = "";
  if (diagnostico) {
    diagnostico.innerHTML = "";
    diagnostico.classList.add("hidden");
  }

  // v8.17.44: todas as telas especiais dividem o mesmo contêiner.
  // Ao trocar de guia, limpamos o conteúdo anterior imediatamente para evitar
  // que a Central de Alertas ou outra guia herde HTML de uma página anterior.
  if (painelPremium) {
    const usaPainelCompartilhado = PAGINAS_PAINEL_COMPARTILHADO_V81744.has(pagina);
    if (usaPainelCompartilhado) {
      const titulo = TITULOS_PAINEL_COMPARTILHADO_V81744[pagina] || "guia";
      painelPremium.innerHTML = `<div class="exec-empty">Carregando ${escapeHtml(titulo)}...</div>`;
      painelPremium.classList.remove("hidden");
    } else {
      painelPremium.innerHTML = "";
      painelPremium.classList.add("hidden");
    }
  }
}


function badge(texto, tipo = "neutral") {
  return `<span class="badge ${tipo}">${texto || ""}</span>`;
}

function saldoBadge(saldo) {
  const s = String(saldo || "").trim();
  if (!s) return "";
  return badge(s, s.startsWith("-") ? "danger" : "success");
}

function statusBadge(status) {
  const s = String(status || "OK").toUpperCase();
  return badge(s, s === "OK" ? "success" : "warning");
}

function tempoParaMinutos(valor) {
  let s = String(valor || "").trim();
  if (!s) return 0;

  const sinal = s.startsWith("-") ? -1 : 1;
  s = s.replace("-", "").trim();

  if (s.includes(":")) {
    const partes = s.split(":");
    const h = parseInt(partes[0] || "0", 10);
    const m = parseInt(partes[1] || "0", 10);
    return sinal * ((isNaN(h) ? 0 : h) * 60 + (isNaN(m) ? 0 : m));
  }

  if (s.includes(",")) {
    const partes = s.split(",");
    const h = parseInt(partes[0] || "0", 10);
    const m = parseInt((partes[1] || "0").slice(0, 2), 10);
    return sinal * ((isNaN(h) ? 0 : h) * 60 + (isNaN(m) ? 0 : m));
  }

  return 0;
}

function paramsConsolidado() {
  const p = new URLSearchParams();

  const cc = qs("cc");
  const nome = qs("nome");
  const turno = qs("turno");
  const tipoInc = qs("tipoInconsistencia");
  const filtroInc = qs("filtroInconsistencia");
  const tipo = qs("tipo");
  const dataEl = qs("data");

  if (cc && cc.value) p.set("cc", cc.value);
  if (nome && nome.value) p.set("nome", nome.value);
  if (turno && turno.value) p.set("turno", turno.value);
  if (tipoInc && tipoInc.value) p.set("tipoInconsistencia", tipoInc.value);
  if (filtroInc && filtroInc.value) p.set("filtroInconsistencia", filtroInc.value);
  if (tipo && tipo.value) p.set("tipo", tipo.value);

  if (datasFiltro.length > 0) {
    p.set("datas", datasFiltro.join(","));
  } else if (dataEl && dataEl.value) {
    p.set("data", dataEl.value);
  }

  return p.toString();
}

function paramsDashboard() {
  const p = new URLSearchParams();

  const turno = qs("turnoDashboard");
  const cc = qs("ccDashboard");
  const nome = qs("nomeDashboard");
  const data = qs("dataDashboard");

  if (turno && turno.value) p.set("turno", turno.value);
  if (cc && cc.value) p.set("cc", cc.value);
  if (nome && nome.value) p.set("nome", nome.value);

  if (datasDashboardFiltro.length > 0) {
    p.set("datas", datasDashboardFiltro.join(","));
  } else if (data && data.value) {
    p.set("data", data.value);
  }

  const ordemSaldo = qs("ordemSaldoBanco");
  if (ordemSaldo && ordemSaldo.value) p.set("order_saldo", ordemSaldo.value);

  return p.toString();
}

function aplicarOrdenacao(rows) {
  let saida = [...rows];

  if (ordenacaoSaldo) {
    const fator = ordenacaoSaldo === "asc" ? 1 : -1;
    saida.sort((a, b) => (tempoParaMinutos(a.saldo) - tempoParaMinutos(b.saldo)) * fator);
  }

  if (ordenacaoHE) {
    const fator = ordenacaoHE === "asc" ? 1 : -1;
    saida.sort((a, b) => (tempoParaMinutos(a.he) - tempoParaMinutos(b.he)) * fator);
  }

  if (ordenacaoInconsistencias) {
    const fator = ordenacaoInconsistencias === "asc" ? 1 : -1;
    saida.sort((a, b) => {
      const av = String(a.inconsistencias || "").toLowerCase();
      const bv = String(b.inconsistencias || "").toLowerCase();
      return av.localeCompare(bv, "pt-BR") * fator;
    });
  }

  return saida;
}

function alternarOrdenacaoSaldo() {
  ordenacaoHE = null;
  ordenacaoInconsistencias = null;

  if (ordenacaoSaldo === null) ordenacaoSaldo = "desc";
  else if (ordenacaoSaldo === "desc") ordenacaoSaldo = "asc";
  else ordenacaoSaldo = null;

  renderConsolidado();
}

function alternarOrdenacaoHE() {
  ordenacaoSaldo = null;
  ordenacaoInconsistencias = null;

  if (ordenacaoHE === null) ordenacaoHE = "desc";
  else if (ordenacaoHE === "desc") ordenacaoHE = "asc";
  else ordenacaoHE = null;

  renderConsolidado();
}

function alternarOrdenacaoInconsistencias() {
  ordenacaoSaldo = null;
  ordenacaoHE = null;
  ordenacaoInconsistencias = null;

  if (ordenacaoInconsistencias === null) ordenacaoInconsistencias = "asc";
  else if (ordenacaoInconsistencias === "asc") ordenacaoInconsistencias = "desc";
  else ordenacaoInconsistencias = null;

  renderConsolidado();
}

function atualizarIndicadoresOrdenacao() {
  const saldoIcon = qs("sortSaldoIcon");
  if (saldoIcon) {
    saldoIcon.textContent =
      ordenacaoSaldo === "desc" ? " ↓" :
      ordenacaoSaldo === "asc" ? " ↑" : " ↕";
  }

  const heIcon = qs("sortHEIcon");
  if (heIcon) {
    heIcon.textContent =
      ordenacaoHE === "desc" ? " ↓" :
      ordenacaoHE === "asc" ? " ↑" : " ↕";
  }

  const incIcon = qs("sortInconsistenciasIcon");
  if (incIcon) {
    incIcon.textContent =
      ordenacaoInconsistencias === "desc" ? " ↓" :
      ordenacaoInconsistencias === "asc" ? " ↑" : " ↕";
  }
}


function atualizarVisibilidadeFiltroTipo() {
  const tipo = qs("tipo");
  if (!tipo) return;

  const deveOcultar = paginaAtual === "hora_extra" || paginaAtual === "ausencias";
  tipo.style.display = deveOcultar ? "none" : "block";
}

function setPaginaConsolidado() {
  ocultarDashboardExecutivoPremium();
  ocultarPainelBancoHoras2();
  ajustarFiltroTipoInconsistencias();
  marcarPaginaAtualClasse();
  ajustarControleOrdenacaoBancoHoras();
  renderizarCardsPorGuia();
  ajustarFiltroDataBancoHoras();
  ajustarCardsBancoHoras();
  const ordemBox = qs("controleOrdenacaoBancoHoras");
  if (ordemBox) ordemBox.classList.add("hidden");

  qs("tabelaConsolidado").classList.remove("hidden");
  qs("tabelaDashboard").classList.add("hidden");
  qs("filtros").classList.remove("hidden");
  qs("filtrosDashboard").classList.add("hidden");
  if (typeof carregarOpcoesFiltros === "function") {
    const selIds = ["cc", "nome", "turno"];
    const precisa = selIds.some(id => qs(id) && qs(id).options.length <= 1);
    if (precisa) carregarOpcoesFiltros(true);
  }

  const tipoInc = qs("tipoInconsistencia");
  if (tipoInc) {
    tipoInc.style.display = paginaAtual === "inconsistencias" ? "block" : "none";
  }

  const filtroIncSelect = qs("filtroInconsistencia");
  if (filtroIncSelect) {
    filtroIncSelect.style.display = paginaAtual === "inconsistencias" ? "block" : "none";
  }

  if (paginaAtual === "inconsistencias") {
    renderizarCardsPorGuia();
    setCabecalhoInconsistenciasCompacto();
  } else if (paginaAtual === "ponto_aberto") {
    setCabecalhoSemMarcacaoCompacto();
  } else if (paginaAtual === "hora_extra") {
    renderizarCardsPorGuia();
    setCabecalhoHoraExtraCompacto();
  } else if (paginaAtual === "ponto_aberto_impar") {
    setCabecalhoPontoAbertoCompacto();
  } else {
    setCabecalhoEspelhoPonto();
  }

  qs("tituloTabela").textContent =
    paginaAtual === "inconsistencias"
      ? "Fila de Inconsistências"
      : "Base Consolidada";

  atualizarVisibilidadeFiltroTipo();
}


function ajustarCardsBancoHoras() {
  const ocultarHEAbsent = paginaAtual === "dashboard";
  const mostrarCardsBanco = paginaAtual === "dashboard";

  const cardHE = qs("cardHE");
  const cardAbsent = qs("cardAbsent");
  const cardBancoPositivo = qs("cardBancoPositivo");
  const cardColabBancoPositivo = qs("cardColabBancoPositivo");
  const cardBancoNegativoTotal = qs("cardBancoNegativoTotal");

  if (cardHE) cardHE.classList.toggle("hidden", ocultarHEAbsent);
  if (cardAbsent) cardAbsent.classList.toggle("hidden", ocultarHEAbsent);

  if (cardBancoPositivo) cardBancoPositivo.classList.toggle("hidden", !mostrarCardsBanco);
  if (cardColabBancoPositivo) cardColabBancoPositivo.classList.toggle("hidden", !mostrarCardsBanco);
  if (cardBancoNegativoTotal) cardBancoNegativoTotal.classList.toggle("hidden", !mostrarCardsBanco);
}



function ajustarFiltroDataBancoHoras() {
  const ocultar = paginaAtual === "dashboard";

  const boxData = qs("boxDataDashboard");
  const boxDatas = qs("boxDatasDashboardSelecionadas");
  const dataInput = qs("dataDashboard");

  if (boxData) boxData.classList.toggle("hidden", ocultar);
  if (boxDatas) boxDatas.classList.toggle("hidden", ocultar);

  // Ao entrar no Banco de Horas, limpa a data para não contaminar o painel.
  if (ocultar && dataInput) dataInput.value = "";

  if (ocultar && typeof datasDashboardFiltro !== "undefined") {
    datasDashboardFiltro = [];
    if (typeof renderDatasDashboardSelecionadas === "function") {
      renderDatasDashboardSelecionadas();
    }
  }
}



function ajustarControleOrdenacaoBancoHoras() {
  const controle = qs("controleOrdenacaoBancoHoras");
  if (!controle) return;

  const mostrar = paginaAtual === "dashboard";
  controle.classList.toggle("hidden", !mostrar);

  if (!mostrar && qs("ordemSaldoBanco")) {
    qs("ordemSaldoBanco").value = "";
  }
}




function atualizarAcoesTopbarPorPagina() {
  const botoesRestritos = [
    "btnTopbarMinhaSenha",
    "btnTopbarSair",
    "btnTopbarImportarExcel",
    "btnTopbarBaixarExcel"
  ];
  const guiasPermitidas = ["administracao", "consolidado"];
  const mostrarBotoesRestritos = guiasPermitidas.includes(paginaAtual);

  botoesRestritos.forEach(id => {
    const el = qs(id);
    if (!el) return;
    el.classList.toggle("hidden", !mostrarBotoesRestritos);
    el.style.display = mostrarBotoesRestritos ? "" : "none";
    if (mostrarBotoesRestritos) {
      el.removeAttribute("aria-hidden");
    } else {
      el.setAttribute("aria-hidden", "true");
    }
  });

  const btnAtualizar = qs("btnTopbarAtualizarDados");
  if (btnAtualizar) {
    const ocultarAtualizar = [
      "administracao",
      "configuracao_robo",
      "admin_config_robo",
      "processamento_integrado",
      "processamento_manual_robo",
      "central_operacao",
      "diagnostico_instalacao",
      "painel_implantacao",
      "auditoria_premium"
    ].includes(paginaAtual);
    btnAtualizar.classList.toggle("hidden", ocultarAtualizar);
    btnAtualizar.style.display = ocultarAtualizar ? "none" : "";
    if (ocultarAtualizar) {
      btnAtualizar.setAttribute("aria-hidden", "true");
    } else {
      btnAtualizar.removeAttribute("aria-hidden");
    }
  }
}

function marcarPaginaAtualClasse() {
  if (!document.body) return;
  document.body.setAttribute("data-pagina-atual", paginaAtual || "");
  document.body.classList.toggle("pagina-dashboard", paginaAtual === "dashboard");
  document.body.classList.toggle("pagina-dashboard-executivo", paginaAtual === "dashboard_executivo");
  document.body.classList.toggle("pagina-alertas-automaticos", paginaAtual === "alertas_automaticos");
  document.body.classList.toggle("pagina-auditoria-premium", paginaAtual === "auditoria_premium");
  document.body.classList.toggle("pagina-score-operacional", paginaAtual === "score_operacional");
  document.body.classList.toggle("pagina-configuracao-robo", paginaAtual === "configuracao_robo");
  document.body.classList.toggle("pagina-processamento-integrado", paginaAtual === "processamento_integrado");
  document.body.classList.toggle("pagina-processamento-manual-robo", paginaAtual === "processamento_manual_robo");
  document.body.classList.toggle("pagina-central-operacao", paginaAtual === "central_operacao");
  document.body.classList.toggle("pagina-diagnostico-instalacao", paginaAtual === "diagnostico_instalacao");
  document.body.classList.toggle("pagina-painel-implantacao", paginaAtual === "painel_implantacao");
  document.body.classList.toggle("pagina-administracao", paginaAtual === "administracao");
  document.body.classList.toggle("pagina-consolidado", paginaAtual === "consolidado");
  atualizarAcoesTopbarPorPagina();
  limparSubtitulosGuiasV82130();
  if (paginaAtual === "alertas_automaticos" && typeof limparSubtituloCentralAlertasV8214 === "function") limparSubtituloCentralAlertasV8214();
}

function setVisibilidadeDashboardExecutivoPremium(mostrar) {
  const painel = qs("dashboardExecutivoPremium");
  const wrap = document.querySelector(".table-wrap");
  if (painel) painel.classList.toggle("hidden", !mostrar);
  if (wrap) wrap.classList.toggle("hidden", mostrar);
}

function ocultarDashboardExecutivoPremium() {
  setVisibilidadeDashboardExecutivoPremium(false);
  const painel = qs("dashboardExecutivoPremium");
  if (painel) painel.classList.add("hidden");
}



const CARDS_LAYOUT_VIOLACOES_JORNADA = `
  <div class="card"><span>Registros</span><strong id="kpiRegistros">0</strong></div>
  <div class="card"><span>Colaboradores</span><strong id="kpiColaboradores">0</strong></div>
  <div class="card"><span>Quebras 12x36</span><strong id="kpiQuebras12x36">0</strong></div>
  <div class="card"><span>Excessos Semanais</span><strong id="kpiExcessosSemanais">0</strong></div>
`;

const CARDS_LAYOUT_PADRAO = `
  <div class="card"><span>Registros</span><strong id="kpiRegistros">0</strong></div>
  <div class="card"><span>Colaboradores</span><strong id="kpiColaboradores">0</strong></div>
  <div class="card"><span>Horas Extras</span><strong id="kpiHE">00:00</strong></div>
  <div class="card"><span>Ausências</span><strong id="kpiAbsent">00:00</strong></div>
  <div class="card"><span>Colab. Saldo Negativo</span><strong id="kpiSaldoNegativo">0</strong></div>
`;

const CARDS_LAYOUT_OPERACIONAL = `
  <div class="card"><span>Registros</span><strong id="kpiRegistros">0</strong></div>
  <div class="card"><span>Colaboradores</span><strong id="kpiColaboradores">0</strong></div>
`;

const CARDS_LAYOUT_HORA_EXTRA = `
  <div class="card"><span>Registros</span><strong id="kpiRegistros">0</strong></div>
  <div class="card"><span>Colaboradores</span><strong id="kpiColaboradores">0</strong></div>
  <div class="card"><span>Horas Extras</span><strong id="kpiHE">00:00</strong></div>
`;

const CARDS_LAYOUT_BANCO_HORAS = `
  <div class="card"><span>Registros</span><strong id="kpiRegistros">0</strong></div>
  <div class="card"><span>Colaboradores</span><strong id="kpiColaboradores">0</strong></div>
  <div class="card"><span>Colab. Saldo Negativo</span><strong id="kpiSaldoNegativo">0</strong></div>
  <div class="card" id="cardBancoNegativoTotal"><span>Banco Negativo Total</span><strong id="kpiBancoNegativoTotal">00,00</strong></div>
  <div class="card" id="cardColabBancoPositivo"><span>Colab. Saldo Positivo</span><strong id="kpiColabBancoPositivo">0</strong></div>
  <div class="card" id="cardBancoPositivo"><span>Banco Positivo</span><strong id="kpiBancoPositivo">00,00</strong></div>
`;


function ajustarFiltroTipoInconsistencias() {
  const boxTipo = qs("boxFiltroTipo");
  if (!boxTipo) return;

  const ocultar = paginaAtual === "inconsistencias";
  boxTipo.classList.toggle("hidden", ocultar);

  if (ocultar && qs("tipo")) {
    qs("tipo").value = "";
  }
}


function renderizarCardsPorGuia() {
  const container = qs("cardsContainer");
  if (!container) return;
  container.classList.remove("hidden");

  let layout = CARDS_LAYOUT_PADRAO;

  if (paginaAtual === "violacoes_jornada") {
    layout = CARDS_LAYOUT_VIOLACOES_JORNADA;
  } else if (paginaAtual === "dashboard") {
    layout = CARDS_LAYOUT_BANCO_HORAS;
  } else if (paginaAtual === "hora_extra") {
    layout = CARDS_LAYOUT_HORA_EXTRA;
  } else if (["inconsistencias", "ponto_aberto", "ponto_aberto_impar", "inter_jornada"].includes(paginaAtual)) {
    layout = CARDS_LAYOUT_OPERACIONAL;
  }

  if (container.dataset.layoutAtual !== paginaAtual) {
    container.innerHTML = layout;
    container.dataset.layoutAtual = paginaAtual;
  }
}

function _htmlAnotacaoBalloon(anotacaoId, temAnotacao) {
  const cls = temAnotacao ? "has-text" : "no-text";
  return `<span class="anotacao-balloon ${cls}" data-anotacao-id="${anotacaoId}" onclick="abrirEditorAnotacao(this)">💬</span>`;
}

function _appendBaloesAposRender(ids) {
  const linhas = document.querySelectorAll("#tbody tr");
  const cabecalho = document.querySelector("#tabelaConsolidado thead tr");
  if (!cabecalho || !linhas.length) return;

  if (!cabecalho.querySelector(".th-anotacao")) {
    const th = document.createElement("th");
    th.className = "th-anotacao";
    th.textContent = "";
    cabecalho.appendChild(th);
  }

  linhas.forEach((tr, i) => {
    if (tr.classList.contains("anotacao-editor-row")) return;
    const id = ids[i] || "";
    const temAnotacao = !!window._anotacoesCache?.[id];
    let td = tr.querySelector("td.td-anotacao");
    if (!td) {
      td = document.createElement("td");
      td.className = "td-anotacao";
      td.style.textAlign = "center";
      tr.appendChild(td);
    }
    td.innerHTML = _htmlAnotacaoBalloon(id, temAnotacao);
  });
}

async function _carregarAnotacoes(ids) {
  if (!ids || !ids.length) return;
  const val_ids = ids.filter(Boolean);
  if (!val_ids.length) return;
  try {
    const resp = await fetch(`/api/anotacoes?ids=${encodeURIComponent(val_ids.join(","))}`);
    const json = await resp.json();
    window._anotacoesCache = window._anotacoesCache || {};
    if (json.anotacoes) Object.assign(window._anotacoesCache, json.anotacoes);
    document.querySelectorAll(".anotacao-balloon").forEach(el => {
      const id = el.dataset.anotacaoId;
      if (window._anotacoesCache?.[id]) {
        el.classList.remove("no-text");
        el.classList.add("has-text");
      } else {
        el.classList.remove("has-text");
        el.classList.add("no-text");
      }
    });
  } catch (_) {}
}

let _editorAberto = null;

function abrirEditorAnotacao(el) {
  if (_editorAberto) fecharEditorAnotacao();
  const id = el.dataset.anotacaoId;
  const tr = el.closest("tr");
  if (!tr) return;
  const textoAtual = (window._anotacoesCache?.[id] || "") + "";

  const editorTr = document.createElement("tr");
  editorTr.className = "anotacao-editor-row";
  editorTr.innerHTML = `
    <td colspan="99">
      <div class="anotacao-editor-wrap">
        <textarea placeholder="Digite sua anotação...">${escapeHtml(textoAtual)}</textarea>
        <div class="anotacao-editor-actions">
          <button class="btn" onclick="fecharEditorAnotacao()">Cancelar</button>
          <button class="btn" onclick="salvarEditorAnotacao()">💾 Salvar</button>
        </div>
      </div>
    </td>`;
  tr.parentNode.insertBefore(editorTr, tr.nextSibling);

  const ta = editorTr.querySelector("textarea");
  ta.focus();
  ta.setSelectionRange(ta.value.length, ta.value.length);

  _editorAberto = { id, tr, editorTr, el };
  ta.addEventListener("keydown", function (e) {
    if (e.key === "Escape") { fecharEditorAnotacao(); e.preventDefault(); }
    if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) { salvarEditorAnotacao(); e.preventDefault(); }
  });
}

function fecharEditorAnotacao() {
  if (_editorAberto) {
    _editorAberto.editorTr.remove();
    _editorAberto = null;
  }
}

async function salvarEditorAnotacao() {
  if (!_editorAberto) return;
  const { id, el } = _editorAberto;
  const textarea = _editorAberto.editorTr.querySelector("textarea");
  const texto = textarea ? textarea.value.trim() : "";

  try {
    const resp = await fetch("/api/anotacoes", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ id, anotacao: texto }),
    });
    const json = await resp.json();
    if (json.ok) {
      window._anotacoesCache = window._anotacoesCache || {};
      window._anotacoesCache[id] = texto;
      el.classList.remove("no-text");
      el.classList.add("has-text");
      fecharEditorAnotacao();
    } else {
      alert("Erro ao salvar anotacao: " + (json.erro || "desconhecido"));
    }
  } catch (e) {
    alert("Erro de rede ao salvar anotacao.");
  }
}

function setPaginaDashboard() {
  ocultarDashboardExecutivoPremium();
  ajustarFiltroTipoInconsistencias();
  marcarPaginaAtualClasse();
  ajustarControleOrdenacaoBancoHoras();
  renderizarCardsPorGuia();
  ajustarFiltroDataBancoHoras();
  ajustarCardsBancoHoras();
  const ordemBox = qs("controleOrdenacaoBancoHoras");
  if (ordemBox) ordemBox.classList.remove("hidden");

  const tipo = qs("tipo");
  if (tipo) tipo.style.display = "block";

  qs("tabelaConsolidado").classList.add("hidden");
  qs("tabelaDashboard").classList.remove("hidden");
  qs("filtros").classList.add("hidden");
  qs("filtrosDashboard").classList.remove("hidden");
  qs("tituloTabela").textContent = "Banco de Horas por Colaborador";
}

function formatarDataBR(dataIso) {
  if (!dataIso || !dataIso.includes("-")) return dataIso || "";
  const [ano, mes, dia] = dataIso.split("-");
  return `${dia}/${mes}/${ano}`;
}

function renderDatasSelecionadas() {
  const box = qs("datasSelecionadas");
  if (!box) return;

  box.innerHTML = datasFiltro.map(data => `
    <span class="date-chip">
      ${formatarDataBR(data)}
      <button type="button" onclick="removerDataFiltro('${data}')">×</button>
    </span>
  `).join("");
}

function adicionarDataFiltro() {
  const input = qs("data");
  if (!input || !input.value) return;

  if (!datasFiltro.includes(input.value)) {
    datasFiltro.push(input.value);
    datasFiltro.sort();
  }

  input.value = "";
  renderDatasSelecionadas();
  recarregarPaginaAtualFiltros();
}

function removerDataFiltro(data) {
  datasFiltro = datasFiltro.filter(d => d !== data);
  renderDatasSelecionadas();
  recarregarPaginaAtualFiltros();
}

function limparDatasFiltro() {
  datasFiltro = [];
  renderDatasSelecionadas();
}

function renderDatasDashboardSelecionadas() {
  const box = qs("datasDashboardSelecionadas");
  if (!box) return;

  box.innerHTML = datasDashboardFiltro.map(data => `
    <span class="date-chip">
      ${formatarDataBR(data)}
      <button type="button" onclick="removerDataDashboardFiltro('${data}')">×</button>
    </span>
  `).join("");
}

function adicionarDataDashboardFiltro() {
  const input = qs("dataDashboard");
  if (!input || !input.value) return;

  if (!datasDashboardFiltro.includes(input.value)) {
    datasDashboardFiltro.push(input.value);
    datasDashboardFiltro.sort();
  }

  input.value = "";
  renderDatasDashboardSelecionadas();
  carregarDashboard();
}

function removerDataDashboardFiltro(data) {
  datasDashboardFiltro = datasDashboardFiltro.filter(d => d !== data);
  renderDatasDashboardSelecionadas();
  carregarDashboard();
}

function limparDatasDashboardFiltro() {
  datasDashboardFiltro = [];
  renderDatasDashboardSelecionadas();
}


function normalizarCCFront(valor) {
  let texto = String(valor || "").trim();
  if (!texto || texto.toLowerCase() === "nan" || texto.toLowerCase() === "none") return "";

  texto = texto.split("-", 1)[0].trim();

  if (texto.endsWith(".0")) {
    texto = texto.slice(0, -2);
  }

  return texto.replace(/\D/g, "");
}

function aplicarFiltroCCFront(rows, selectId) {
  const select = qs(selectId);
  if (!select || !select.value) return rows;

  const ccSelecionado = normalizarCCFront(select.value);

  return (rows || []).filter(r => {
    const ccLinha = normalizarCCFront(r.cc);
    return ccLinha === ccSelecionado;
  });
}

async function carregarConsolidado() {
  const paginaEsperada = paginaAtual;
  const tokenEsperado = tokenAtualPagina();
  setPaginaConsolidado();
  qs("statusCarga").textContent = "Carregando...";

  const queryConsolidado = paramsConsolidado();
  console.log("Query Consolidado:", queryConsolidado);
  const { controller, timer } = criarControllerNavegacao(TIMEOUT_CONSULTA_ESPELHO_MS);
  let resp;
  let json;
  try {
    resp = await fetch(`/api/consolidado?${queryConsolidado}`, { signal: controller.signal });
    json = await resp.json();
  } catch (e) {
    if (!requisicaoAindaValida(paginaEsperada, tokenEsperado)) return;
    qs("statusCarga").textContent = e.name === "AbortError" ? "Consulta ainda está pesada e foi interrompida. Aplique filtro de data/CC/nome ou recarregue após a primeira indexação do Excel." : "Erro ao carregar dados: " + e.message;
    return;
  } finally {
    finalizarControllerNavegacao(controller, timer);
  }
  if (!requisicaoAindaValida(paginaEsperada, tokenEsperado)) return;

  if (json.erro) {
    qs("statusCarga").textContent = json.erro;
    qs("tbody").innerHTML = "";
    zerarKpis();
    dadosConsolidadoAtual = [];
    return;
  }

  const k = json.kpis || {};
  if (qs("kpiRegistros")) qs("kpiRegistros").textContent = k.registros ?? 0;
  if (qs("kpiColaboradores")) qs("kpiColaboradores").textContent = k.colaboradores ?? 0;
  if (qs("kpiHE")) qs("kpiHE").textContent = k.he_total ?? "00:00";
  if (qs("kpiAbsent")) qs("kpiAbsent").textContent = k.absent_total ?? "00:00";
  if (qs("kpiSaldoNegativo")) qs("kpiSaldoNegativo").textContent = k.saldo_negativo ?? 0;
  if (qs("kpiQuebras12x36")) qs("kpiQuebras12x36").textContent = k.quebras_12x36 ?? 0;
  if (qs("kpiExcessosSemanais")) qs("kpiExcessosSemanais").textContent = k.excessos_semanais ?? 0;
  if (qs("kpiRevisar")) qs("kpiRevisar").textContent = k.revisar ?? 0;

  dadosConsolidadoAtual = aplicarFiltroCCFront(json.dados || [], "cc");

  if (paginaAtual === "inconsistencias") {
    dadosConsolidadoAtual = dadosConsolidadoAtual.filter(r => {
      const status = String(r.status || "").toUpperCase();
      const inc = String(r.inconsistencias || "").trim();
      return status !== "OK" || inc !== "";
    });

    const tipoInc = String(qs("tipoInconsistencia")?.value || "").toLowerCase().trim();
    const filtroInc = String(qs("filtroInconsistencia")?.value || "").toLowerCase().trim();

    if (tipoInc) {
      dadosConsolidadoAtual = dadosConsolidadoAtual.filter(r =>
        String(r.inconsistencias || "").toLowerCase().includes(tipoInc)
      );
    }

    if (filtroInc) {
      dadosConsolidadoAtual = dadosConsolidadoAtual.filter(r => {
        const partes = String(r.inconsistencias || "")
          .split(";")
          .map(p => p.trim().toLowerCase())
          .filter(Boolean);
        const textoCompleto = String(r.inconsistencias || "").trim().toLowerCase();
        return partes.includes(filtroInc) || textoCompleto === filtroInc;
      });
    }
  }

  renderConsolidado();
  const idsCons = (dadosConsolidadoAtual || []).map(r => r._anotacao_id);
  _appendBaloesAposRender(idsCons);
  _carregarAnotacoes(idsCons);
  if (json.limitado && qs("statusCarga")) {
    qs("statusCarga").textContent += ` Exibição limitada a ${json.limite_linhas} de ${json.total_sem_limite} registro(s). Use filtros para refinar.`;
  }
}


function renderInconsistenciasCompacta(rows) {
  qs("tbody").innerHTML = limitarRenderTabela(rows, "inconsistências").map(r => `
    <tr>
      <td>${r.cc ?? ""}</td>
      <td><strong>${r.nome ?? ""}</strong></td>
      <td>${r.data ?? ""}</td>
      <td>${r.dia ?? ""}</td>

      <td>${r.e1 ?? ""}</td>
      <td>${r.s1 ?? ""}</td>
      <td>${r.e2 ?? ""}</td>
      <td>${r.s2 ?? ""}</td>
      <td>${r.e3 ?? ""}</td>
      <td>${r.s3 ?? ""}</td>
      <td>${r.e4 ?? ""}</td>
      <td>${r.s4 ?? ""}</td>

      <td>${r.jornada ?? ""}</td>
      <td>${r.observacao ?? ""}</td>
      <td>${typeof statusBadge === 'function' ? statusBadge(r.status) : (r.status ?? "")}</td>
      <td>${r.inconsistencias ?? ""}</td>
    </tr>
  `).join("");
}



function setCabecalhoInterJornada() {
  const thead = document.querySelector("#tabelaConsolidado thead tr");
  if (!thead) return;

  thead.innerHTML = `
    <th>CC</th>
    <th>Nome</th>
    <th>Função</th>
    <th>Turno</th>
    <th>Data Anterior</th>
    <th>Última Marcação</th>
    <th>Data Retorno</th>
    <th>Primeira Marcação</th>
    <th>Intervalo</th>
    <th>Déficit</th>
    <th>Observação</th>
  `;
}





function renderInterJornada(rows) {
  const linhas = limitarRenderTabela(rows, "interjornadas");
  const tbody = qs("tbody");
  if (!tbody) return;

  tbody.innerHTML = linhas.map(r => `
    <tr>
      <td>${r.cc ?? ""}</td>
      <td><strong>${r.nome ?? ""}</strong></td>
      <td>${r.funcao ?? ""}</td>
      <td>${r.turno ?? ""}</td>
      <td>${r.data_anterior ?? ""}</td>
      <td>${r.ultima_marcacao_anterior ?? ""}</td>
      <td>${r.data_retorno ?? ""}</td>
      <td>${r.primeira_marcacao_retorno ?? ""}</td>
      <td>${r.intervalo_interjornada ?? ""}</td>
      <td>${r.deficit_interjornada ?? ""}</td>
      <td>${r.observacao ?? ""}</td>
    </tr>
  `).join("");

  if (qs("statusCarga")) qs("statusCarga").textContent = `${linhas.length} ocorrência(s) de violações inter jornada carregada(s).`;
}





async function carregarInterJornada() {
  const paginaEsperada = "inter_jornada";
  const tokenEsperado = tokenAtualPagina();
  paginaAtual = "inter_jornada";
  dadosConsolidadoAtual = [];

  if (typeof renderizarCardsPorGuia === "function") renderizarCardsPorGuia();

  if (qs("tituloPagina")) qs("tituloPagina").textContent = "Violações Inter Jornada";
  if (qs("subtituloPagina")) qs("subtituloPagina").textContent = "Ocorrências em que o intervalo entre a última marcação e a primeira marcação seguinte é inferior a 11 horas.";
  if (qs("tituloTabela")) qs("tituloTabela").textContent = "Violações Inter Jornada";
  if (qs("statusCarga")) qs("statusCarga").textContent = "Carregando ocorrências de violações inter jornada...";
  if (qs("tbody")) qs("tbody").innerHTML = "";

  setCabecalhoInterJornada();

  try {
    const resp = await fetch(`/api/inter-jornada?${paramsConsolidado()}`);
    const json = await resp.json();
    if (!requisicaoAindaValida(paginaEsperada, tokenEsperado)) return;

    if (!resp.ok || json.erro) {
      if (qs("statusCarga")) qs("statusCarga").textContent = json.erro || `Erro HTTP ${resp.status}`;
      if (qs("tbody")) qs("tbody").innerHTML = "";
      if (typeof zerarKpis === "function") zerarKpis();
      return;
    }

    const k = json.kpis || {};
    if (qs("kpiRegistros")) qs("kpiRegistros").textContent = k.registros ?? 0;
    if (qs("kpiColaboradores")) qs("kpiColaboradores").textContent = k.colaboradores ?? 0;

    dadosConsolidadoAtual = json.dados || [];
    setCabecalhoInterJornada();
    renderInterJornada(dadosConsolidadoAtual);
    const idsIJ = (dadosConsolidadoAtual || []).map(r => r._anotacao_id);
    _appendBaloesAposRender(idsIJ);
    _carregarAnotacoes(idsIJ);
    if (typeof renderizarCardsPorGuia === "function") renderizarCardsPorGuia();
  } catch (e) {
    if (qs("statusCarga")) qs("statusCarga").textContent = "Erro ao carregar Violações Inter Jornada: " + e.message;
    if (qs("tbody")) qs("tbody").innerHTML = "";
  }
}






function setCabecalhoEspelhoPonto() {
  const thead = document.querySelector("#tabelaConsolidado thead tr");
  if (!thead) return;

  // v8.10.56: Espelho de Ponto mais leve e aderente.
  // As colunas Saldo Atual, Status e Inconsistências foram suprimidas
  // desta guia para evitar leitura/uso de dados auxiliares que não representam
  // o objetivo principal do Espelho. As guias específicas continuam preservadas.
  thead.innerHTML = `
    <th>CC</th><th>Nome</th><th>Função</th><th>Turno</th><th>Data</th><th>Dia</th>
    <th>1a E.</th><th>1a S.</th><th>2a E.</th><th>2a S.</th><th>3a E.</th><th>3a S.</th><th>4a E.</th><th>4a S.</th>
    <th>Abono</th>
    <th><button class="th-sort" onclick="alternarOrdenacaoHE()">H.E.<span id="sortHEIcon"> ↕</span></button></th>
    <th>Absent.</th><th>Jornada</th><th>Ad. Not.</th><th>Observação</th>
  `;
}

function setCabecalhoInconsistenciasCompacto() {
  const thead = document.querySelector("#tabelaConsolidado thead tr");
  if (!thead) return;

  thead.innerHTML = `
    <th>CC</th>
    <th>Nome</th>
    <th>Data</th>
    <th>Dia</th>
    <th>1a E.</th>
    <th>1a S.</th>
    <th>2a E.</th>
    <th>2a S.</th>
    <th>3a E.</th>
    <th>3a S.</th>
    <th>4a E.</th>
    <th>4a S.</th>
    <th>Jornada</th>
    <th>Observação</th>
    <th>Status</th>
    <th><button class="th-sort" onclick="alternarOrdenacaoInconsistencias()">Inconsistências<span id="sortInconsistenciasIcon"> ↕</span></button></th>
  `;
}

function renderInconsistenciasCompacta(rows) {
  const linhas = aplicarOrdenacao(rows || []);

  qs("tbody").innerHTML = linhas.map(r => `
    <tr>
      <td>${r.cc ?? ""}</td>
      <td><strong>${r.nome ?? ""}</strong></td>
      <td>${r.data ?? ""}</td>
      <td>${r.dia ?? ""}</td>
      <td>${r.e1 ?? ""}</td>
      <td>${r.s1 ?? ""}</td>
      <td>${r.e2 ?? ""}</td>
      <td>${r.s2 ?? ""}</td>
      <td>${r.e3 ?? ""}</td>
      <td>${r.s3 ?? ""}</td>
      <td>${r.e4 ?? ""}</td>
      <td>${r.s4 ?? ""}</td>
      <td>${r.jornada ?? ""}</td>
      <td>${r.observacao ?? ""}</td>
      <td>${statusBadge(r.status)}</td>
      <td>${r.inconsistencias ?? ""}</td>
    </tr>
  `).join("");

  atualizarIndicadoresOrdenacao();

  qs("statusCarga").textContent = `${linhas.length} inconsistência(s) carregada(s).`;
}


function setCabecalhoSemMarcacaoCompacto() {
  const thead = document.querySelector("#tabelaConsolidado thead tr");
  if (!thead) return;

  thead.innerHTML = `
    <th>CC</th>
    <th>Nome</th>
    <th>Função</th>
    <th>Turno</th>
    <th>Data</th>
    <th>Dia</th>
    <th>1a E.</th>
    <th>1a S.</th>
    <th>2a E.</th>
    <th>2a S.</th>
    <th>3a E.</th>
    <th>3a S.</th>
    <th>4a E.</th>
    <th>4a S.</th>
    <th>Abono</th>
    <th>Absent.</th>
    <th>Observação</th>
  `;
}

function renderSemMarcacaoCompacta(rows) {
  const linhas = rows || [];

  qs("tbody").innerHTML = linhas.map(r => `
    <tr>
      <td>${r.cc ?? ""}</td>
      <td><strong>${r.nome ?? ""}</strong></td>
      <td>${r.funcao ?? ""}</td>
      <td>${r.turno ?? ""}</td>
      <td>${r.data ?? ""}</td>
      <td>${r.dia ?? ""}</td>
      <td>${r.e1 ?? ""}</td>
      <td>${r.s1 ?? ""}</td>
      <td>${r.e2 ?? ""}</td>
      <td>${r.s2 ?? ""}</td>
      <td>${r.e3 ?? ""}</td>
      <td>${r.s3 ?? ""}</td>
      <td>${r.e4 ?? ""}</td>
      <td>${r.s4 ?? ""}</td>
      <td>${r.abono ?? ""}</td>
      <td>${r.absent ? badge(r.absent, "danger") : ""}</td>
      <td>${r.observacao ?? ""}</td>
    </tr>
  `).join("");

  qs("statusCarga").textContent = `${linhas.length} registro(s) sem marcação carregado(s).`;
}


function setCabecalhoPontoAbertoCompacto() {
  const thead = document.querySelector("#tabelaConsolidado thead tr");
  if (!thead) return;

  thead.innerHTML = `
    <th>CC</th>
    <th>Nome</th>
    <th>Função</th>
    <th>Turno</th>
    <th>Data</th>
    <th>Dia</th>
    <th>1a E.</th>
    <th>1a S.</th>
    <th>2a E.</th>
    <th>2a S.</th>
    <th>3a E.</th>
    <th>3a S.</th>
    <th>4a E.</th>
    <th>4a S.</th>
    <th>Absent.</th>
    <th>Observação</th>
    <th>Status</th>
    <th>Inconsistências</th>
  `;
}

function renderPontoAbertoCompacto(rows) {
  const linhas = somenteImparesPontoAberto(rows || []);

  qs("tbody").innerHTML = linhas.map(r => `
    <tr>
      <td>${r.cc ?? ""}</td>
      <td><strong>${r.nome ?? ""}</strong></td>
      <td>${r.funcao ?? ""}</td>
      <td>${r.turno ?? ""}</td>
      <td>${r.data ?? ""}</td>
      <td>${r.dia ?? ""}</td>
      <td>${r.e1 ?? ""}</td>
      <td>${r.s1 ?? ""}</td>
      <td>${r.e2 ?? ""}</td>
      <td>${r.s2 ?? ""}</td>
      <td>${r.e3 ?? ""}</td>
      <td>${r.s3 ?? ""}</td>
      <td>${r.e4 ?? ""}</td>
      <td>${r.s4 ?? ""}</td>
      <td>${r.absent ? badge(r.absent, "danger") : ""}</td>
      <td>${r.observacao ?? ""}</td>
      <td>${statusBadge(r.status)}</td>
      <td>${r.inconsistencias ?? ""}</td>
    </tr>
  `).join("");

  const nomes = new Set(linhas.map(r => String(r.nome || "").trim()).filter(Boolean));

  if (qs("kpiRegistros")) qs("kpiRegistros").textContent = linhas.length;
  if (qs("kpiColaboradores")) qs("kpiColaboradores").textContent = nomes.size;
  if (qs("kpiHE")) qs("kpiHE").textContent = "00:00";
  if (qs("kpiAbsent")) qs("kpiAbsent").textContent = "00:00";
  if (qs("kpiSaldoNegativo")) qs("kpiSaldoNegativo").textContent = "0";
  if (qs("kpiRevisar")) qs("kpiRevisar").textContent = linhas.length;

  qs("statusCarga").textContent = `${linhas.length} registro(s) com marcações ímpares carregado(s).`;
}







function setCabecalhoHoraExtraCompacto() {
  const thead = document.querySelector("#tabelaConsolidado thead tr");
  if (!thead) return;

  thead.innerHTML = `
    <th>CC</th>
    <th>Nome</th>
    <th>Função</th>
    <th>Turno</th>
    <th>Data</th>
    <th>Dia</th>
    <th>1a E.</th>
    <th>1a S.</th>
    <th>2a E.</th>
    <th>2a S.</th>
    <th>3a E.</th>
    <th>3a S.</th>
    <th>4a E.</th>
    <th>4a S.</th>
    <th><button class="th-sort" onclick="alternarOrdenacaoHE()">H.E.<span id="sortHEIcon"> ↕</span></button></th>
    <th>Jornada</th>
    <th>Observação</th>
    <th>Status</th>
    <th>Inconsistências</th>
  `;
}

function renderHoraExtraCompacta(rows) {
  const linhas = aplicarOrdenacao(rows || []);

  qs("tbody").innerHTML = linhas.map(r => `
    <tr>
      <td>${r.cc ?? ""}</td>
      <td><strong>${r.nome ?? ""}</strong></td>
      <td>${r.funcao ?? ""}</td>
      <td>${r.turno ?? ""}</td>
      <td>${r.data ?? ""}</td>
      <td>${r.dia ?? ""}</td>
      <td>${r.e1 ?? ""}</td>
      <td>${r.s1 ?? ""}</td>
      <td>${r.e2 ?? ""}</td>
      <td>${r.s2 ?? ""}</td>
      <td>${r.e3 ?? ""}</td>
      <td>${r.s3 ?? ""}</td>
      <td>${r.e4 ?? ""}</td>
      <td>${r.s4 ?? ""}</td>
      <td>${r.he ? badge(r.he, "warning") : ""}</td>
      <td>${r.jornada ?? ""}</td>
      <td>${r.observacao ?? ""}</td>
      <td>${statusBadge(r.status)}</td>
      <td>${r.inconsistencias ?? ""}</td>
    </tr>
  `).join("");

  atualizarIndicadoresOrdenacao();

  qs("statusCarga").textContent = `${linhas.length} registro(s) de hora extra carregado(s).`;
}

function renderConsolidado() {
  if (paginaAtual === "hora_extra") {
    setCabecalhoHoraExtraCompacto();
    renderHoraExtraCompacta(dadosConsolidadoAtual);
    return;
  }


  if (paginaAtual === "inconsistencias") {
    setCabecalhoInconsistenciasCompacto();
    renderInconsistenciasCompacta(dadosConsolidadoAtual);
    return;
  } else {
    setCabecalhoEspelhoPonto();
  }


  const rows = limitarRenderTabela(aplicarOrdenacao(dadosConsolidadoAtual), "registros");

  const renderLinha = (r) => `
    <tr>
      <td>${r.cc}</td>
      <td><strong>${r.nome}</strong></td>
      <td>${r.funcao}</td>
      <td>${r.turno}</td>
      <td>${r.data}</td>
      <td>${r.dia}</td>
      <td>${r.e1}</td>
      <td>${r.s1}</td>
      <td>${r.e2}</td>
      <td>${r.s2}</td>
      <td>${r.e3}</td>
      <td>${r.s3}</td>
      <td>${r.e4}</td>
      <td>${r.s4}</td>
      <td>${r.abono}</td>
      <td>${r.he ? badge(r.he, "warning") : ""}</td>
      <td>${r.absent ? badge(r.absent, "danger") : ""}</td>
      <td>${r.jornada}</td>
      <td>${r.adnot}</td>
      <td>${r.observacao}</td>
    </tr>
  `;

  atualizarIndicadoresOrdenacao();

  let msg = `${rows.length} registro(s) carregado(s).`;
  if (ordenacaoHE === "desc") msg += " H.E. ordenada do maior para o menor.";
  if (ordenacaoHE === "asc") msg += " H.E. ordenada do menor para o maior.";

  const paginaRenderizada = paginaAtual;
  qs("statusCarga").textContent = rows.length ? `Renderizando ${rows.length} registro(s)...` : "Nenhum registro encontrado.";
  renderTabelaEmLotes("tbody", rows, renderLinha, {
    aoConcluir: () => {
      if (paginaAtual === paginaRenderizada) qs("statusCarga").textContent = msg;
    },
  });
}


async function carregarPontoEmAberto() {
  ajustarFiltroTipoInconsistencias();
  paginaAtual = "ponto_aberto";
    renderizarCardsPorGuia();
  renderizarCardsPorGuia();
  paginaAtual = "ponto_aberto";
  renderizarCardsPorGuia();
  paginaAtual = "ponto_aberto";
  setPaginaConsolidado();
  qs("tituloTabela").textContent = "Sem Marcação";
  qs("statusCarga").textContent = "Carregando registros sem marcação...";

  const resp = await fetch(`/api/ponto-em-aberto?${paramsConsolidado()}`);
  const json = await resp.json();

  if (json.erro) {
    qs("statusCarga").textContent = json.erro;
    qs("tbody").innerHTML = "";
    zerarKpis();
    dadosConsolidadoAtual = [];
    return;
  }

  const k = json.kpis || {};
  if (qs("kpiRegistros")) qs("kpiRegistros").textContent = k.registros ?? 0;
  if (qs("kpiColaboradores")) qs("kpiColaboradores").textContent = k.colaboradores ?? 0;
  if (qs("kpiHE")) qs("kpiHE").textContent = k.he_total ?? "00:00";
  if (qs("kpiAbsent")) qs("kpiAbsent").textContent = k.absent_total ?? "00:00";
  if (qs("kpiSaldoNegativo")) qs("kpiSaldoNegativo").textContent = k.saldo_negativo ?? 0;
  if (qs("kpiRevisar")) qs("kpiRevisar").textContent = k.revisar ?? 0;

  dadosConsolidadoAtual = json.dados || [];
  setCabecalhoSemMarcacaoCompacto();
  renderSemMarcacaoCompacta(dadosConsolidadoAtual);
  const idsPontoAberto = (dadosConsolidadoAtual || []).map(r => r._anotacao_id);
  _appendBaloesAposRender(idsPontoAberto);
  _carregarAnotacoes(idsPontoAberto);
  renderizarCardsPorGuia();
}


function qtdBatidasFront(row) {
  const bruto = row.qtd_batidas ?? row["qtd_batidas"] ?? "";
  const n = parseInt(String(bruto).replace(",", ".").trim(), 10);
  return Number.isNaN(n) ? 0 : n;
}

function filtrarPontoAbertoFront(rows) {
  return (rows || []).filter(r => {
    const qtd = qtdBatidasFront(r);
    return qtd > 0 && qtd % 2 === 1;
  });
}


function qtdBatidasPontoAberto(row) {
  const bruto = row.qtd_batidas ?? row["qtd_batidas"] ?? row["Qtd Batidas"] ?? "";
  const texto = String(bruto || "").trim().replace(",", ".");
  const n = parseInt(texto, 10);
  return Number.isNaN(n) ? 0 : n;
}

function somenteImparesPontoAberto(rows) {
  return (rows || []).filter(r => {
    const qtd = qtdBatidasPontoAberto(r);
    return qtd > 0 && qtd % 2 === 1;
  });
}

function renderPontoAberto(rows) {
  rows = somenteImparesPontoAberto(rows || []);

  qs("tbody").innerHTML = rows.map(r => `
    <tr>
      <td>${r.cc}</td><td><strong>${r.nome}</strong></td><td>${r.funcao}</td><td>${r.turno}</td>
      <td>${r.data}</td><td>${r.dia}</td>
      <td>${r.e1}</td><td>${r.s1}</td><td>${r.e2}</td><td>${r.s2}</td><td>${r.e3}</td><td>${r.s3}</td><td>${r.e4}</td><td>${r.s4}</td>
      <td>${r.abono}</td>
      <td>${r.he ? badge(r.he, "warning") : ""}</td>
      <td>${r.absent ? badge(r.absent, "danger") : ""}</td>
      <td>${r.jornada}</td><td>${r.adnot}</td><td>${r.observacao}</td>
      <td>${saldoBadge(r.saldo)}</td><td>${statusBadge(r.status)}</td><td>${r.inconsistencias}</td>
    </tr>
  `).join("");

  const nomes = new Set(rows.map(r => String(r.nome || "").trim()).filter(Boolean));

  if (qs("kpiRegistros")) qs("kpiRegistros").textContent = rows.length;
  if (qs("kpiColaboradores")) qs("kpiColaboradores").textContent = nomes.size;
  if (qs("kpiHE")) qs("kpiHE").textContent = "00:00";
  if (qs("kpiAbsent")) qs("kpiAbsent").textContent = "00:00";
  if (qs("kpiSaldoNegativo")) qs("kpiSaldoNegativo").textContent = "0";
  if (qs("kpiRevisar")) qs("kpiRevisar").textContent = rows.length;

  qs("statusCarga").textContent = `${rows.length} registro(s) com marcações ímpares.`;
}

async function carregarPontoAberto() {
  ajustarFiltroTipoInconsistencias();
  paginaAtual = "ponto_aberto_impar";
    renderizarCardsPorGuia();
  renderizarCardsPorGuia();
  paginaAtual = "ponto_aberto_impar";
  renderizarCardsPorGuia();
  paginaAtual = "ponto_aberto_impar";
  setPaginaConsolidado();
  qs("tituloTabela").textContent = "Ponto Aberto";
  qs("statusCarga").textContent = "Carregando registros com marcações ímpares...";

  const resp = await fetch(`/api/ponto-aberto?${paramsConsolidado()}`);
  const json = await resp.json();

  if (json.erro) {
    qs("statusCarga").textContent = json.erro;
    qs("tbody").innerHTML = "";
    zerarKpis();
    dadosConsolidadoAtual = [];
    return;
  }

  dadosConsolidadoAtual = somenteImparesPontoAberto(json.dados || []);
  setCabecalhoPontoAbertoCompacto();
  renderPontoAbertoCompacto(dadosConsolidadoAtual);
  const idsPontoAbertoImpar = (dadosConsolidadoAtual || []).map(r => r._anotacao_id);
  _appendBaloesAposRender(idsPontoAbertoImpar);
  _carregarAnotacoes(idsPontoAbertoImpar);
  renderizarCardsPorGuia();
}




function saldoBancoParaMinutosSelect(valor) {
  const textoOriginal = String(valor || "").trim();
  if (!textoOriginal) return 0;

  const negativo = textoOriginal.startsWith("-");
  let limpo = textoOriginal.replace("+", "").replace("-", "").trim();

  // Aceita saldo como HH,MM ou HH:MM.
  // Ex.: 210,26 = 210h26min | -45,30 = -45h30min
  limpo = limpo.replace(",", ":");

  const partes = limpo.split(":");
  if (partes.length < 2) return 0;

  const horas = parseInt(partes[0], 10) || 0;
  const minutos = parseInt(partes[1], 10) || 0;
  const total = (horas * 60) + minutos;

  return negativo ? -total : total;
}

function aplicarOrdenacaoSaldoBancoSelect(rows) {
  const seletor = qs("ordemSaldoBanco");
  const ordem = seletor ? seletor.value : "";
  const dados = [...(rows || [])];

  if (!ordem) return dados;

  dados.sort((a, b) => {
    const av = saldoBancoParaMinutosSelect(a.saldo_atual ?? a.saldo_total ?? a.saldo ?? "");
    const bv = saldoBancoParaMinutosSelect(b.saldo_atual ?? b.saldo_total ?? b.saldo ?? "");
    return ordem === "asc" ? av - bv : bv - av;
  });

  return dados;
}



function saldoBancoCardParaMinutos(valor) {
  const textoOriginal = String(valor || "").trim();
  if (!textoOriginal) return 0;

  const negativo = textoOriginal.startsWith("-");
  let limpo = textoOriginal.replace("+", "").replace("-", "").trim();
  limpo = limpo.replace(",", ":");

  const partes = limpo.split(":");
  if (partes.length < 2) return 0;

  const horas = parseInt(partes[0], 10) || 0;
  const minutos = parseInt(partes[1], 10) || 0;
  const total = (horas * 60) + minutos;

  return negativo ? -total : total;
}

function minutosParaSaldoBancoCard(totalMinutos) {
  const negativo = totalMinutos < 0;
  const abs = Math.abs(totalMinutos);
  const h = Math.floor(abs / 60);
  const m = abs % 60;
  return `${negativo ? "-" : ""}${String(h).padStart(2, "0")},${String(m).padStart(2, "0")}`;
}

function calcularBancoPositivo(rows) {
  return (rows || []).reduce((acc, r) => {
    const saldoMin = saldoBancoCardParaMinutos(r.saldo_atual ?? r.saldo_total ?? r.saldo ?? "");
    return saldoMin > 0 ? acc + saldoMin : acc;
  }, 0);
}


function calcularMetricasBancoHoras(rows) {
  const metricas = {
    positivoMin: 0,
    negativoMinAbs: 0,
    colabPositivo: 0,
    colabNegativo: 0
  };

  (rows || []).forEach(r => {
    const saldoMin = saldoBancoCardParaMinutos(r.saldo_atual ?? r.saldo_total ?? r.saldo ?? "");

    if (saldoMin > 0) {
      metricas.positivoMin += saldoMin;
      metricas.colabPositivo += 1;
    } else if (saldoMin < 0) {
      metricas.negativoMinAbs += Math.abs(saldoMin);
      metricas.colabNegativo += 1;
    }
  });

  return metricas;
}




let bh2ResumoConfig = { total_colaboradores: 0, total_opcoes: 0 };
let bh2RenderSeq = 0;

function colaboradorSelecionadoBH2() {
  // v8.15.15: BH 2.0 pertence somente a guia Banco de Horas.
  // Nunca le o filtro #nome da guia Extrato/Hora Extra para evitar residuos
  // de selecao entre guias.
  const el = qs('nomeDashboard');
  return el ? String(el.value || '').trim() : '';
}

function labelOrigemPadraoBH2(valor) {
  const mapa = {
    saldo_atual: 'Saldo Atual',
    manual: 'Valor Manual'
  };
  return mapa[valor] || valor || 'Saldo Atual';
}


function ocultarPainelBancoHoras2() {
  bh2RenderSeq += 1;
  const painel = qs("painelBancoHoras2");
  if (!painel) return;
  painel.classList.add("hidden");
  painel.innerHTML = "";
  painel.removeAttribute("data-colaborador-bh2");
}

function limparResiduosBancoHoras2ForaDaGuiaBanco() {
  // v8.15.16: isolamento definitivo. O painel/configuracao BH 2.0
  // pertence apenas a guia Banco de Horas (paginaAtual === "dashboard").
  // Extrato, Hora Extra e demais guias nunca devem herdar colaborador, saldo
  // manual, origem de saldo ou qualquer HTML residual do painel BH2.
  if (paginaAtual !== "dashboard") {
    ocultarPainelBancoHoras2();
  }
}

function alternarSaldoManualBH2() {
  const origem = qs('bh2OrigemColaborador') ? qs('bh2OrigemColaborador').value : 'saldo_atual';
  const box = qs('bh2SaldoManualBox');
  const input = qs('bh2SaldoManual');
  const dica = qs('bh2SaldoManualDica');
  if (!box || !input) return;

  // v8.15.13: o campo deve permanecer visivel quando houver colaborador
  // selecionado. Ele fica apenas desabilitado quando a origem for Saldo Atual,
  // evitando a sensacao de que a opcao manual desapareceu da tela.
  box.classList.remove('hidden');
  const nomeSelecionado = colaboradorSelecionadoBH2();
  if (!nomeSelecionado) {
    input.disabled = true;
    input.placeholder = 'Selecione um colaborador';
    if (dica) dica.textContent = 'Selecione um colaborador para habilitar o saldo inicial manual.';
    return;
  }
  if (origem === 'manual') {
    input.disabled = false;
    input.placeholder = 'Ex.: 144,08';
    if (dica) dica.textContent = 'Informe o saldo inicial apenas para o colaborador selecionado.';
  } else {
    input.disabled = true;
    if (!input.value) input.placeholder = 'Habilitado ao escolher Manual';
    if (dica) dica.textContent = 'Campo habilitado somente quando a origem for Manual.';
  }
}

async function renderizarPainelBancoHoras2() {
  const painel = qs("painelBancoHoras2");
  if (!painel) return;

  const seq = ++bh2RenderSeq;
  const paginaOrigem = paginaAtual;

  if (paginaOrigem !== "dashboard") {
    ocultarPainelBancoHoras2();
    return;
  }
  painel.classList.remove("hidden");

  const nomeSelecionado = colaboradorSelecionadoBH2();
  const params = new URLSearchParams();
  if (nomeSelecionado) params.set('nome', nomeSelecionado);

  painel.innerHTML = ``;
  painel.classList.add("hidden");
  try {
    const resp = await fetch(`/api/banco-horas-2/config?${params.toString()}`);
    const json = await resp.json();

    // Se o usuario trocou de guia enquanto a chamada estava em andamento,
    // nao renderiza o painel BH2 em Hora Extra, Extrato ou qualquer outra guia.
    if (seq !== bh2RenderSeq || paginaAtual !== "dashboard") {
      ocultarPainelBancoHoras2();
      return;
    }
    const cfg = json.config || {};
    const cfgColab = json.config_colaborador || {};
    const dataInicio = cfg.data_inicio || '';
    const origemColab = cfgColab.origem || 'saldo_atual';
    bh2ResumoConfig = { total_colaboradores: json.total_colaboradores || 0, total_opcoes: json.total_opcoes || 0 };

    // v8.19.5: restaura a escolha visivel entre saldo inicial automatico e manual
    // na guia Banco de Horas. O controle fica sempre presente; quando nenhum
    // colaborador esta selecionado, ele orienta o usuario em vez de desaparecer.
    const origemSelecionada = origemColab === 'manual' ? 'manual' : 'saldo_atual';
    const painelColaborador = `
      <div class="filters inline-filters bh2-topo bh2-saldo-inicial">
        <label>Saldo inicial
          <select id="bh2OrigemColaborador" onchange="alternarSaldoManualBH2()" ${nomeSelecionado ? '' : 'disabled'}>
            <option value="saldo_atual" ${origemSelecionada === 'saldo_atual' ? 'selected' : ''}>Automático</option>
            <option value="manual" ${origemSelecionada === 'manual' ? 'selected' : ''}>Manual</option>
          </select>
        </label>
        <label id="bh2SaldoManualBox">Valor manual
          <input id="bh2SaldoManual" placeholder="${nomeSelecionado ? (origemSelecionada === 'manual' ? 'Ex.: 144,08' : 'Habilitado ao escolher Manual') : 'Selecione um colaborador'}" value="${escapeHtml(cfgColab.saldo || '')}" ${nomeSelecionado && origemSelecionada === 'manual' ? '' : 'disabled'}>
        </label>
        <button id="bh2BtnSalvar" type="button" class="btn" onclick="salvarBancoHoras2()" ${nomeSelecionado ? '' : 'disabled'}>Aplicar saldo</button>
        <span id="bh2Status" class="diag-note">${nomeSelecionado ? 'Escolha Automático ou Manual para o colaborador selecionado.' : 'Selecione um colaborador para habilitar o saldo inicial manual.'}</span>
      </div>
      `;

    renderizarControlesDataInicialBH2(dataInicio);
    painel.classList.remove("hidden");
    painel.innerHTML = `<div class="exec-panel bh2-config bh2-por-colaborador bh2-compacto">
      ${json.erro ? `<p class="warn">${escapeHtml(json.erro)}</p>` : ''}
      ${painelColaborador}
    </div>`;
    alternarSaldoManualBH2();
  } catch (e) {
    painel.innerHTML = `<div class="exec-panel bh2-config compact"><p class="erro">Falha ao carregar configuração: ${escapeHtml(e.message || e)}</p></div>`;
  }
}


async function salvarDataInicialBancoHoras2() {
  const dataInicio = qs('bh2DataInicio') ? qs('bh2DataInicio').value : '';
  const btn = qs('bh2BtnSalvarData');
  const status = qs('bh2StatusData');
  try {
    if (btn) { btn.disabled = true; btn.textContent = 'Atualizando...'; }
    if (status) status.textContent = dataInicio ? 'Salvando data inicial...' : 'Limpando data inicial para considerar todo o período...';
    const resp = await fetch('/api/banco-horas-2/config', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ ativo: true, data_inicio: dataInicio })
    });
    let json = {};
    try { json = await resp.json(); } catch (parseErr) { json = { ok: false, erro: 'Resposta inválida do servidor.' }; }
    if (!resp.ok || !json.ok) {
      throw new Error(json.erro || json.mensagem || `Falha HTTP ${resp.status}`);
    }
    if (status) status.textContent = json.mensagem || (dataInicio ? 'Data inicial atualizada.' : 'Data inicial limpa. Todo o período será considerado.');
    if (paginaAtual === 'dashboard' && typeof carregarDashboard === 'function') {
      carregarDashboard();
    } else {
      await renderizarPainelBancoHoras2();
    }
  } catch (e) {
    const msg = e && e.message ? e.message : String(e);
    if (status) status.textContent = `Falha ao atualizar data: ${msg}`;
    alert(`Falha ao atualizar data inicial: ${msg}`);
  } finally {
    if (btn) { btn.disabled = false; btn.textContent = 'Atualizar data inicial'; }
  }
}

async function salvarBancoHoras2() {
  const nomeSelecionado = colaboradorSelecionadoBH2();
  const origem = qs('bh2OrigemColaborador') ? qs('bh2OrigemColaborador').value : 'saldo_atual';
  const btn = qs('bh2BtnSalvar');
  const status = qs('bh2Status');
  const payload = {
    ativo: true,
    data_inicio: (qs('bh2DataInicio') ? qs('bh2DataInicio').value : ''),
    nome: nomeSelecionado,
    origem: origem,
    saldo_manual: (qs('bh2SaldoManual') ? qs('bh2SaldoManual').value.trim() : '')
  };
  if (!payload.nome) {
    alert('Selecione um colaborador antes de configurar o saldo inicial.');
    return;
  }
  if (payload.origem === 'manual' && !payload.saldo_manual) {
    alert('Informe o saldo manual no formato HH,MM para este colaborador.');
    return;
  }
  try {
    if (btn) { btn.disabled = true; btn.textContent = 'Atualizando...'; }
    if (status) status.textContent = 'Salvando regra do colaborador e recalculando cache...';
    const resp = await fetch('/api/banco-horas-2/config', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload) });
    let json = {};
    try { json = await resp.json(); } catch (parseErr) { json = { ok: false, erro: 'Resposta inválida do servidor.' }; }
    if (!resp.ok || !json.ok) {
      throw new Error(json.erro || json.mensagem || `Falha HTTP ${resp.status}`);
    }
    if (status) status.textContent = json.mensagem || 'Banco de Horas atualizado.';
    alert(json.mensagem || 'Banco de Horas atualizado.');
    // v8.15.12: salvar regra do Banco não deve disparar a guia Extrato.
    // Cada guia recarrega apenas seus próprios dados para evitar estado cruzado.
    if (paginaAtual === 'dashboard' && typeof carregarDashboard === 'function') {
      carregarDashboard();
    } else {
      await renderizarPainelBancoHoras2();
    }
  } catch (e) {
    const msg = e && e.message ? e.message : String(e);
    if (status) status.textContent = `Falha ao atualizar: ${msg}`;
    alert(`Falha ao atualizar Banco de Horas: ${msg}`);
  } finally {
    if (btn) { btn.disabled = false; btn.textContent = 'Aplicar para este colaborador'; }
  }
}

function recarregarPainelBancoHoras2() { renderizarPainelBancoHoras2(); }

function renderizarControlesDataInicialBH2(dataInicio) {
  const status = qs("statusCarga");
  if (!status || paginaAtual !== "dashboard") return;
  status.classList.add("bh2-header-controls");
  status.innerHTML = `
    <label>Data inicial <input id="bh2DataInicio" type="date" value="${escapeHtml(dataInicio || '')}"></label>
    <button id="bh2BtnSalvarData" type="button" class="btn secondary" onclick="salvarDataInicialBancoHoras2()">Atualizar data</button>
    <button class="btn secondary" onclick="recarregarPainelBancoHoras2()">Recarregar</button>
    <span id="bh2StatusData" class="diag-note"></span>
  `;
}

function limparControlesDataInicialBH2() {
  const status = qs("statusCarga");
  if (!status) return;
  status.classList.remove("bh2-header-controls");
}

async function carregarDashboard() {
  const paginaEsperada = "dashboard";
  const tokenEsperado = tokenAtualPagina();
  ajustarFiltroTipoInconsistencias();
  marcarPaginaAtualClasse();
  paginaAtual = "dashboard";
  ajustarControleOrdenacaoBancoHoras();
  paginaAtual = "dashboard";
  renderizarCardsPorGuia();
  paginaAtual = "dashboard";
  ajustarFiltroDataBancoHoras();
  paginaAtual = "dashboard";
  ajustarCardsBancoHoras();
  setPaginaDashboard();
  if (qs("statusCarga")) qs("statusCarga").textContent = "Carregando tabela principal do Banco de Horas...";
  if (qs("tbodyDashboard")) {
    qs("tbodyDashboard").innerHTML = `<tr><td colspan="9">Consulta principal em execucao. A tabela sera preenchida assim que o resumo responder.</td></tr>`;
  }
  let avisoConsultaBancoHorasTimer = setTimeout(() => {
    if (paginaAtual !== "dashboard") return;
    if (qs("statusCarga")) qs("statusCarga").textContent = "Banco de Horas ainda em processamento. A consulta principal continua em execucao...";
    if (qs("tbodyDashboard")) {
      qs("tbodyDashboard").innerHTML = `<tr><td colspan="9">Processo em execucao no servidor. Aguarde mais alguns instantes; filtros e painel auxiliar foram adiados para priorizar esta tabela.</td></tr>`;
    }
  }, 900);

  // v8.21.12: a tabela principal tem prioridade. Filtros e painel BH2 entram
  // depois para nao competir com a consulta que preenche a guia.
  const carregarComplementosBancoHoras = () => {
    carregarOpcoesFiltrosBancoHoras().catch(e => {
      console.warn("Falha ao carregar filtros do Banco de Horas em segundo plano.", e);
    });
    renderizarPainelBancoHoras2().catch(e => {
      console.warn("Falha ao carregar painel auxiliar do Banco de Horas em segundo plano.", e);
    });
  };

  const { controller, timer } = criarControllerNavegacao(TIMEOUT_CONSULTA_ESPELHO_MS);
  let json;
  try {
    const resp = await fetch(`/api/dashboard-banco-he?${paramsDashboard()}`, { signal: controller.signal });
    json = await resp.json();
  } catch (e) {
    clearTimeout(avisoConsultaBancoHorasTimer);
    if (!requisicaoAindaValida(paginaEsperada, tokenEsperado)) return;
    qs("statusCarga").textContent = e.name === "AbortError" ? "Consulta do Banco de Horas interrompida. Use filtros para reduzir o volume ou tente novamente." : "Erro ao carregar Banco de Horas: " + e.message;
    return;
  } finally {
    finalizarControllerNavegacao(controller, timer);
  }
  clearTimeout(avisoConsultaBancoHorasTimer);
  if (!requisicaoAindaValida(paginaEsperada, tokenEsperado)) return;

  if (json.erro) {
    qs("statusCarga").textContent = json.erro;
    qs("tbodyDashboard").innerHTML = "";
    return;
  }

  const k = json.kpis || {};
  if (qs("kpiRegistros")) qs("kpiRegistros").textContent = k.registros ?? 0;
  if (qs("kpiColaboradores")) qs("kpiColaboradores").textContent = k.colaboradores ?? 0;
  if (qs("kpiHE")) qs("kpiHE").textContent = k.he_total ?? "00:00";
  if (qs("kpiAbsent")) qs("kpiAbsent").textContent = k.absent_total ?? "00:00";
  if (qs("kpiSaldoNegativo")) qs("kpiSaldoNegativo").textContent = k.saldo_negativo ?? 0;
  if (qs("kpiRevisar")) qs("kpiRevisar").textContent = k.revisar ?? 0;

  const dadosDashboardFiltrados = aplicarFiltroCCFront(json.dados || [], "ccDashboard");

  const bancoPositivoMin = calcularBancoPositivo(dadosDashboardFiltrados);
  if (qs("kpiBancoPositivo")) qs("kpiBancoPositivo").textContent = minutosParaSaldoBancoCard(bancoPositivoMin);

  const metricasBanco = calcularMetricasBancoHoras(dadosDashboardFiltrados);
  if (qs("kpiBancoPositivo")) qs("kpiBancoPositivo").textContent = minutosParaSaldoBancoCard(metricasBanco.positivoMin);
  if (qs("kpiColabBancoPositivo")) qs("kpiColabBancoPositivo").textContent = metricasBanco.colabPositivo;
  if (qs("kpiBancoNegativoTotal")) qs("kpiBancoNegativoTotal").textContent = "-" + minutosParaSaldoBancoCard(metricasBanco.negativoMinAbs);
  if (qs("kpiSaldoNegativo")) qs("kpiSaldoNegativo").textContent = metricasBanco.colabNegativo;

  // v8.21.9: Dias com HE e Dias com ausência agora são métricas reais.
  // Se algum cache antigo não trouxer a métrica, exibimos "-" em vez de
  // inventar zero, para não transformar ausência de dado em número.
  const valorDiasBancoHoras = (valor) => {
    if (valor === null || valor === undefined) return "-";
    const txt = String(valor).trim();
    if (!txt || txt.toLowerCase() === "undefined" || txt.toLowerCase() === "nan" || txt.toLowerCase() === "none") return "-";
    return escapeHtml(txt);
  };

  const renderLinhaDashboard = (r) => `
    <tr>
      <td>${escapeHtml(r.cc ?? "")}</td>
      <td><strong>${escapeHtml(r.nome ?? "")}</strong></td>
      <td>${escapeHtml(r.funcao ?? "")}</td>
      <td>${escapeHtml(r.turno ?? "")}</td>
      <td class="hide-banco-horas-col">${badge(r.he_total || "00:00", "warning")}</td>
      <td>${valorDiasBancoHoras(r.dias_com_he)}</td>
      <td class="hide-banco-horas-col">${badge(r.absent_total || "00:00", "danger")}</td>
      <td>${valorDiasBancoHoras(r.dias_com_ausencia)}</td>
      <td>${saldoBadge(r.saldo_atual)}</td>
    </tr>
  `;

  renderTabelaEmLotes("tbodyDashboard", dadosDashboardFiltrados, renderLinhaDashboard, {
    htmlVazio: `<tr><td colspan="9">${escapeHtml(json.aviso || "Nenhum colaborador carregado no resumo do Banco de Horas.")}</td></tr>`,
    aoConcluir: (total) => {
      if (paginaAtual === "dashboard" && qs("statusCarga") && !qs("statusCarga").classList.contains("bh2-header-controls")) {
        qs("statusCarga").textContent = json.aviso || `${total} colaborador(es) carregado(s) no Banco de Horas.`;
      }
    },
  });
  setTimeout(() => {
    if (paginaAtual === "dashboard") carregarComplementosBancoHoras();
  }, 120);

  // v8.15.27: a guia Banco de Horas usa o canto do titulo para os controles
  // da Data Inicial. Mensagens tecnicas como fonte/cache ficam ocultas para
  // preservar espaco util da tela.
  if (qs("statusCarga") && !qs("statusCarga").classList.contains("bh2-header-controls")) {
    qs("statusCarga").textContent = json.aviso || `Renderizando ${dadosDashboardFiltrados.length} colaborador(es) do Banco de Horas...`;
  }
}


async function carregarOpcoesInconsistencias() {
  const select = qs("filtroInconsistencia");
  if (!select) return;

  const valorAtual = select.value;

  try {
    const resp = await fetch("/api/inconsistencias-opcoes");
    const json = await resp.json();

    select.innerHTML = '<option value="">Todas as inconsistências</option>';

    (json.opcoes || []).forEach(opcao => {
      const opt = document.createElement("option");
      opt.value = opcao;
      opt.textContent = opcao;
      select.appendChild(opt);
    });

    if (valorAtual) select.value = valorAtual;
  } catch (e) {
    console.warn("Falha ao carregar opções de inconsistências", e);
  }
}



async function carregarUsuarioAtual() {
  const badge = qs("usuarioAtualBadge");
  if (!badge) return;

  try {
    const resp = await fetch("/api/usuario-atual");
    if (resp.status === 401) {
      window.location.href = "/login";
      return;
    }
    const json = await resp.json();
    window.USUARIO_ATUAL = json;
    badge.textContent = `${json.nome || json.usuario} · ${json.perfil || ""}`;
    document.body.classList.toggle("perfil-admin", !!json.is_admin);
    document.body.classList.toggle("perfil-consulta", !json.is_admin);
    aplicarPermissoesInterface();
    if (json.trocar_senha) {
      setTimeout(() => abrirTrocaSenhaRapida(), 500);
    }
  } catch (e) {
    badge.textContent = "Usuário";
  }
}

function usuarioAtualEhAdmin() {
  return !!(window.USUARIO_ATUAL && window.USUARIO_ATUAL.is_admin);
}

function aplicarPermissoesInterface() {
  const isAdmin = usuarioAtualEhAdmin();
  document.querySelectorAll(".requires-admin").forEach(el => {
    el.classList.toggle("admin-hidden", !isAdmin);
    if (!isAdmin) el.setAttribute("aria-hidden", "true");
    else el.removeAttribute("aria-hidden");
  });
}



function vincularEventoComboCC() {
  const cc = qs("cc");
  if (cc && !cc.dataset.ccListener) {
    cc.addEventListener("change", atualizarNomesPorCCPrincipal);
    cc.dataset.ccListener = "1";
  }

  const ccDashboard = qs("ccDashboard");
  if (ccDashboard && !ccDashboard.dataset.ccListener) {
    ccDashboard.addEventListener("change", atualizarNomesPorCCDashboard);
    ccDashboard.dataset.ccListener = "1";
  }
}


async function carregarOpcoesFiltros(force = false) {
  try {
    if (force) limparCacheOpcoesFiltros();
    const json = await obterOpcoesFiltros("", force);
    const ccs = json.ccs || [];
    const nomes = json.nomes || [];
    const turnos = json.turnos || [];

    preencherSelect("cc", ccs, "Todos os CCs");
    preencherSelect("ccDashboard", ccs, "Todos os CCs");
    preencherSelect("nome", nomes, "Todos os colaboradores");
    preencherSelect("nomeDashboard", nomes, "Todos os colaboradores");
    preencherSelect("turno", turnos, "Todos os turnos");
    preencherSelect("turnoDashboard", turnos, "Todos os turnos");

    vincularEventoComboCC();

  } catch (e) {
    console.warn("Falha ao carregar opções dos filtros.", e);
  }
}



async function buscarOpcoesPorCC(cc) {
  const json = await obterOpcoesFiltros(cc || "");
  return { nomes: json.nomes || [], turnos: json.turnos || [] };
}
async function buscarNomesPorCC(cc) {
  const query = cc ? `?cc=${encodeURIComponent(cc)}` : "";
  const resp = await fetch(`/api/nomes-por-cc${query}`);

  if (resp.status === 401) {
    window.location.href = "/login";
    return [];
  }

  const json = await resp.json();
  return json.nomes || [];
}

async function atualizarNomesPorCCPrincipal() {
  const cc = qs("cc") ? qs("cc").value : "";
  const opcoes = await buscarOpcoesPorCC(cc);

  preencherSelect("nome", opcoes.nomes || [], "Todos os colaboradores");
  preencherSelect("turno", opcoes.turnos || [], "Todos os turnos");

  if (qs("nome")) qs("nome").value = "";
  if (qs("turno")) qs("turno").value = "";

  // v8.11.27: ao trocar CC dentro de guias especiais, recarrega a guia atual
  // em vez de jogar o usuário de volta para o Consolidado.
  recarregarPaginaAtualFiltros();
  setTimeout(() => {
    if (paginaAtual === "inconsistencias" && typeof renderInconsistenciasCompacta === 'function' && typeof dadosConsolidadoAtual !== 'undefined') {
      renderInconsistenciasCompacta(dadosConsolidadoAtual);
    }
  }, 50);
}

async function atualizarNomesPorCCDashboard() {
  const cc = qs("ccDashboard") ? qs("ccDashboard").value : "";
  const opcoes = await buscarOpcoesPorCC(cc);

  preencherSelect("nomeDashboard", opcoes.nomes || [], "Todos os colaboradores");
  preencherSelect("turnoDashboard", opcoes.turnos || [], "Todos os turnos");

  if (qs("nomeDashboard")) qs("nomeDashboard").value = "";
  if (qs("turnoDashboard")) qs("turnoDashboard").value = "";

  if (paginaAtual === "dashboard") carregarDashboard();
}


function preencherSelect(id, valores, rotuloTodos) {
  const select = qs(id);
  if (!select) return;

  const valorAtual = select.value;
  select.innerHTML = "";

  const optTodos = document.createElement("option");
  optTodos.value = "";
  optTodos.textContent = rotuloTodos;
  select.appendChild(optTodos);

  valores.forEach(item => {
    const opt = document.createElement("option");
    if (typeof item === "object" && item !== null) {
      opt.value = item.valor || item.rotulo || "";
      opt.textContent = item.rotulo || item.valor || "";
    } else {
      opt.value = item;
      opt.textContent = item;
    }
    select.appendChild(opt);
  });

  if (valorAtual) select.value = valorAtual;
}


async function carregarOpcoesFiltrosBancoHoras(force = false) {
  // v8.15.12: a guia Banco de Horas não depende mais da guia Extrato
  // para popular colaboradores. Ela carrega seus próprios filtros
  // dashboard (#ccDashboard, #nomeDashboard, #turnoDashboard).
  const nomeSelect = qs("nomeDashboard");
  const turnoSelect = qs("turnoDashboard");
  const ccSelect = qs("ccDashboard");
  if (!nomeSelect && !turnoSelect && !ccSelect) return;

  const precisaCarregar = force
    || (nomeSelect && nomeSelect.options.length <= 1)
    || (turnoSelect && turnoSelect.options.length <= 1)
    || (ccSelect && ccSelect.options.length <= 1);

  if (!precisaCarregar) return;

  try {
    const json = await obterOpcoesFiltros(ccSelect ? ccSelect.value : "", force);
    if (ccSelect && (!ccSelect.value || ccSelect.options.length <= 1)) {
      preencherSelect("ccDashboard", json.ccs || [], "Todos os CCs");
    }
    preencherSelect("nomeDashboard", json.nomes || [], "Todos os colaboradores");
    preencherSelect("turnoDashboard", json.turnos || [], "Todos os turnos");
  } catch (e) {
    console.warn("Falha ao carregar filtros independentes do Banco de Horas.", e);
  }
}

async function importarExcelRobo() {
  if (!usuarioAtualEhAdmin()) {
    alert("Acesso negado. A importação do Excel exige perfil Administrador.");
    return;
  }
  const status = qs("statusCarga");
  if (status) status.textContent = "Importando Excel do Robô em modo rápido...";

  try {
    const resp = await fetch("/api/importar-excel-robo", { method: "POST" });
    const json = await resp.json();

    if (!resp.ok || json.erro || json.ok === false) {
      const msg = json.erro || json.mensagem || `Erro HTTP ${resp.status} ao importar Excel.`;
      if (status) status.textContent = msg;
      alert(msg);
      return;
    }

    const msg = json.mensagem || "Excel importado com sucesso.";
    let msgCorrecao = "";
    if (json.correcao_saldo_pdf && json.correcao_saldo_pdf.mensagem) {
      const c = json.correcao_saldo_pdf;
      msgCorrecao = `\n\nBanco de Horas: ${c.mensagem}`;
      if (c.pdfs_lidos !== undefined) msgCorrecao += `\nPDFs/quadros lidos: ${c.pdfs_lidos}`;
      if (c.linhas_corrigidas !== undefined) msgCorrecao += ` | Linhas corrigidas: ${c.linhas_corrigidas}`;
      if (c.segundos !== undefined) msgCorrecao += ` | Tempo: ${c.segundos}s`;
    }
    let msgNotificacao = "";
    if (json.notificacoes_jornada) {
      const n = json.notificacoes_jornada;
      const envios = (n.envios || []).map(e => `${e.canal}: ${e.mensagem}`).join(" | ");
      msgNotificacao = `\n\nAlertas de jornada: ${n.novas ?? 0} nova(s) ocorrência(s). ${envios}`;
    }
    if (status) status.textContent = msg + " Importação concluída. A base será carregada sob demanda ao abrir as guias." + (msgCorrecao ? " Banco de horas ficará disponível sob demanda." : "") + (msgNotificacao ? " Alertas adiados." : "");
    alert(msg + msgCorrecao + msgNotificacao);
    limparCacheOpcoesFiltros();

    // v8.11.21: não recarregar automaticamente Consolidado/Dashboard após a
    // importação. Em bases grandes isso renderizava milhares de linhas e
    // bloqueava a página. Atualizamos filtros em segundo plano e deixamos o
    // usuário escolher a guia a abrir.
    // v8.11.22: não recarrega filtros automaticamente após importar.
    // A tela fica livre imediatamente; os filtros são carregados sob demanda
    // quando a guia escolhida pelo usuário for aberta.

    if (qs("tabelaConsolidado")) qs("tabelaConsolidado").classList.add("hidden");
    if (qs("tabelaDashboard")) qs("tabelaDashboard").classList.add("hidden");
    if (qs("dashboardExecutivoPremium")) qs("dashboardExecutivoPremium").classList.add("hidden");
    if (status) status.textContent = "Excel importado em modo leve. Abra a guia desejada; os dados serão carregados sob demanda.";
  } catch (e) {
    const msg = "Erro ao importar Excel: " + e.message;
    if (status) status.textContent = msg;
    alert(msg);
  }
}




function setCabecalhoViolacoesJornada() {
  const thead = document.querySelector("#tabelaConsolidado thead tr");
  if (!thead) return;

  thead.innerHTML = `
    <th>CC</th>
    <th>Nome</th>
    <th>Função</th>
    <th>Tipo Jornada</th>
    <th>Período</th>
    <th>Data Referência</th>
    <th>Violação</th>
    <th>Detalhe</th>
    <th>Dias Trabalhados</th>
    <th>Limite Esperado</th>
  `;
}





function renderViolacoesJornada(rows) {
  const linhas = limitarRenderTabela(rows, "violações");
  const tbody = qs("tbody");
  if (!tbody) return;

  tbody.innerHTML = linhas.map(r => `
    <tr>
      <td>${r.cc ?? ""}</td>
      <td><strong>${r.nome ?? ""}</strong></td>
      <td>${r.funcao ?? ""}</td>
      <td>${r.tipo_jornada ?? ""}</td>
      <td>${r.periodo ?? ""}</td>
      <td>${r.data_referencia ?? ""}</td>
      <td>${r.violacao ?? ""}</td>
      <td>${r.detalhe ?? ""}</td>
      <td>${r.dias_trabalhados ?? ""}</td>
      <td>${r.limite_esperado ?? ""}</td>
    </tr>
  `).join("");

  if (qs("statusCarga")) qs("statusCarga").textContent = `${linhas.length} violação(ões) de jornada carregada(s).`;
}






async function carregarDiagnosticoViolacoesJornadaSeVazio(totalLinhas) {
  if (totalLinhas && totalLinhas > 0) return;

  try {
    const resp = await fetch("/api/debug-violacoes-jornada");
    const diag = await resp.json();

    if (!diag.erro) {
      const totalTrab = diag.linhas_trabalhadas_detectadas ?? 0;
      const exc = (diag.excessos_semanais_amostra || []).length;
      const cons = (diag.consecutivos_12x36_amostra || []).length;

      qs("statusCarga").textContent =
        `Nenhuma violação exibida. Diagnóstico atualizado: ${totalTrab} dia(s) trabalhado(s) detectado(s), ` +
        `${exc} excesso(s) semanal(is) em amostra, ${cons} sequência(s) 12x36 em amostra.`;
    }
  } catch (e) {}
}


async function carregarViolacoesJornada() {
  const paginaEsperada = "violacoes_jornada";
  const tokenEsperado = tokenAtualPagina();
  paginaAtual = "violacoes_jornada";
  dadosConsolidadoAtual = [];

  if (typeof renderizarCardsPorGuia === "function") renderizarCardsPorGuia();

  if (qs("tituloPagina")) qs("tituloPagina").textContent = "Violações de Jornada";
  if (qs("subtituloPagina")) qs("subtituloPagina").textContent = "Alertas de excesso de jornada conforme regra nominal do turno.";
  if (qs("tituloTabela")) qs("tituloTabela").textContent = "Violações de Jornada";
  if (qs("statusCarga")) qs("statusCarga").textContent = "Carregando violações de jornada...";
  if (qs("tbody")) qs("tbody").innerHTML = "";

  setCabecalhoViolacoesJornada();

  try {
    const resp = await fetch(`/api/violacoes-jornada?${paramsConsolidado()}`);
    const json = await resp.json();
    if (!requisicaoAindaValida(paginaEsperada, tokenEsperado)) return;

    if (!resp.ok || json.erro) {
      if (qs("statusCarga")) qs("statusCarga").textContent = json.erro || `Erro HTTP ${resp.status}`;
      if (qs("tbody")) qs("tbody").innerHTML = "";
      if (typeof zerarKpis === "function") zerarKpis();
      return;
    }

    const k = json.kpis || {};
    if (qs("kpiRegistros")) qs("kpiRegistros").textContent = k.registros ?? 0;
    if (qs("kpiColaboradores")) qs("kpiColaboradores").textContent = k.colaboradores ?? 0;
    if (qs("kpiQuebras12x36")) qs("kpiQuebras12x36").textContent = k.quebras_12x36 ?? 0;
    if (qs("kpiExcessosSemanais")) qs("kpiExcessosSemanais").textContent = k.excessos_semanais ?? 0;

    dadosConsolidadoAtual = json.dados || [];
    setCabecalhoViolacoesJornada();
    renderViolacoesJornada(dadosConsolidadoAtual);
    const idsVJ = (dadosConsolidadoAtual || []).map(r => r._anotacao_id);
    _appendBaloesAposRender(idsVJ);
    _carregarAnotacoes(idsVJ);
    if (typeof renderizarCardsPorGuia === "function") renderizarCardsPorGuia();
  } catch (e) {
    if (qs("statusCarga")) qs("statusCarga").textContent = "Erro ao carregar Violações de Jornada: " + e.message;
    if (qs("tbody")) qs("tbody").innerHTML = "";
  }
}


























































function ocultarFiltroDataPrincipalPremium() {
  const dataInput = qs("data");
  const dataControl = dataInput ? dataInput.closest(".multi-date-control") : null;
  if (dataControl) dataControl.classList.add("hidden");
  if (dataInput) dataInput.value = "";

  const datasSelecionadas = qs("datasSelecionadas");
  if (datasSelecionadas) datasSelecionadas.classList.add("hidden");

  if (typeof datasFiltro !== "undefined") {
    datasFiltro = [];
    if (typeof renderDatasSelecionadas === "function") renderDatasSelecionadas();
  }
}

function ajustarFiltrosExtratoBancoHoras() {
  const dataInput = qs("data");
  const dataControl = dataInput ? dataInput.closest(".multi-date-control") : null;

  // v8.11.28: no Extrato Banco de Horas, os filtros de data também são
  // relevantes. A versão anterior escondia o controle de datas, então o
  // usuário não conseguia conferir visualmente nem alterar o recorte aplicado.
  if (dataControl) dataControl.classList.remove("hidden");

  const datasSelecionadas = qs("datasSelecionadas");
  if (datasSelecionadas) datasSelecionadas.classList.remove("hidden");

  if (qs("boxFiltroTipo")) qs("boxFiltroTipo").classList.add("hidden");
  if (qs("tipo")) qs("tipo").classList.add("hidden");
  if (qs("filtroInconsistencia")) qs("filtroInconsistencia").style.display = "none";
  if (qs("controleOrdenacaoBancoHoras")) qs("controleOrdenacaoBancoHoras").classList.add("hidden");

  document.body.classList.add("pagina-extrato-banco");
}

function restaurarFiltrosPadraoPrincipal() {
  const dataInput = qs("data");
  const dataControl = dataInput ? dataInput.closest(".multi-date-control") : null;
  const paginaPremiumSemData = paginaAtual === "dashboard_executivo" || paginaAtual === "alertas_automaticos" || paginaAtual === "auditoria_premium";
  if (dataControl) dataControl.classList.toggle("hidden", paginaPremiumSemData);

  const datasSelecionadas = qs("datasSelecionadas");
  if (datasSelecionadas) datasSelecionadas.classList.toggle("hidden", paginaPremiumSemData);

  if (qs("boxFiltroTipo")) qs("boxFiltroTipo").classList.remove("hidden");
  if (qs("tipo")) qs("tipo").classList.remove("hidden");
  if (qs("controleOrdenacaoBancoHoras")) qs("controleOrdenacaoBancoHoras").classList.add("hidden");

  document.body.classList.remove("pagina-extrato-banco");
}

async function abrirExtratoBancoHoras(el) {
  iniciarTransicaoPagina("extrato_banco_horas");
  ocultarDashboardExecutivoPremium();
  paginaAtual = "extrato_banco_horas";
  limparResiduosBancoHoras2ForaDaGuiaBanco();
  dadosConsolidadoAtual = [];

  document.querySelectorAll(".menu button").forEach(b => b.classList.remove("active"));
  if (el) el.classList.add("active");

  if (qs("filtros")) qs("filtros").classList.remove("hidden");
  if (qs("filtrosDashboard")) qs("filtrosDashboard").classList.add("hidden");
  if (qs("boxFiltroTipo")) qs("boxFiltroTipo").classList.add("hidden");
  ajustarFiltrosExtratoBancoHoras();
  if (qs("controleOrdenacaoBancoHoras")) qs("controleOrdenacaoBancoHoras").classList.add("hidden");
  if (qs("tabelaDashboard")) qs("tabelaDashboard").classList.add("hidden");
  if (qs("tabelaConsolidado")) qs("tabelaConsolidado").classList.remove("hidden");

  await carregarOpcoesExtratoBancoHoras();
  carregarExtratoBancoHoras();
}

function setCabecalhoExtratoBancoHoras() {
  const thead = document.querySelector("#tabelaConsolidado thead tr");
  if (!thead) return;

  thead.innerHTML = `
    <th>Data</th>
    <th>Data Emissão PDF</th>
    <th>Colaborador</th>
    <th>CC</th>
    <th>Turno</th>
    <th>Dia</th>
    <th>Evento</th>
    <th>Crédito H.E.</th>
    <th>Débito Ausência</th>
    <th>Saldo Calculado</th>
    <th>Saldo Informado</th>
    <th>Observação</th>
    <th>Inconsistências</th>
  `;
}

function renderExtratoBancoHoras(rows) {
  const linhas = rows || [];
  const tbody = qs("tbody");
  if (!tbody) return;

  tbody.innerHTML = linhas.map(r => `
    <tr>
      <td>${r.data ?? ""}</td>
      <td>${r.data_emissao_pdf ?? ""}</td>
      <td><strong>${r.nome ?? ""}</strong></td>
      <td>${r.cc ?? ""}</td>
      <td>${r.turno ?? ""}</td>
      <td>${r.dia ?? ""}</td>
      <td><strong>${r.evento ?? ""}</strong></td>
      <td>${r.he ?? ""}</td>
      <td>${r.absent ?? ""}</td>
      <td><strong>${r.saldo_calculado ?? ""}</strong></td>
      <td>${r.saldo_informado ?? r.saldo_pdf ?? ""}</td>
      <td>${r.observacao ?? ""}</td>
      <td>${r.inconsistencias ?? ""}</td>
    </tr>
  `).join("");

  if (qs("statusCarga")) qs("statusCarga").textContent = `${linhas.length} linha(s) de extrato carregada(s).`;
}

function renderCardsExtratoBancoHoras(resumo) {
  const container = qs("cardsContainer");
  if (!container) return;

  const r = resumo || {};
  const temResumo = !!(r.nome || r.saldo_base || r.saldo_final || r.data_base);

  if (!temResumo) {
    container.innerHTML = `
      <div class="card"><span>Saldo Base</span><strong>—</strong></div>
      <div class="card"><span>Créditos</span><strong>—</strong></div>
      <div class="card"><span>Débitos</span><strong>—</strong></div>
      <div class="card"><span>Saldo Final</span><strong>—</strong></div>
      <div class="card"><span>Data Base</span><strong>—</strong></div>
      <div class="card"><span>Origem</span><strong>—</strong></div>
    `;
    container.dataset.layoutAtual = "extrato_banco_horas_vazio";
    return;
  }

  const origem = r.origem_saldo_base ? r.origem_saldo_base : "base do ciclo";

  container.innerHTML = `
    <div class="card"><span>Saldo Base</span><strong>${r.saldo_base ?? "00,00"}</strong></div>
    <div class="card"><span>Créditos</span><strong>${r.credito_total ?? "00:00"}</strong></div>
    <div class="card"><span>Débitos</span><strong>${r.debito_total ?? "00:00"}</strong></div>
    <div class="card"><span>Saldo Final</span><strong>${r.saldo_final ?? "00,00"}</strong></div>
    <div class="card"><span>Data Base</span><strong>${r.data_base ?? ""}</strong></div>
    <div class="card"><span>Origem</span><strong>${origem}</strong></div>
  `;
  container.dataset.layoutAtual = "extrato_banco_horas";
}


function garantirPainelDiagnosticoBancoHoras() {
  // v8.15.17: painel removido da interface operacional. Mantido como stub seguro.
  let painel = document.getElementById("painelDiagnosticoBancoHoras");
  if (!painel) {
    painel = document.createElement("div");
    painel.id = "painelDiagnosticoBancoHoras";
    painel.className = "diagnostico-banco-horas hidden";
  }
  painel.innerHTML = "";
  painel.classList.add("hidden");
  return painel;
}


function renderObjetoDiagnostico(obj) {
  if (!obj || typeof obj !== "object") return "<em>Sem dados.</em>";
  const linhas = Object.entries(obj).map(([k, v]) => {
    const valor = (v && typeof v === "object") ? `<pre>${escapeHtml(JSON.stringify(v, null, 2))}</pre>` : escapeHtml(String(v ?? ""));
    return `<tr><th>${escapeHtml(k)}</th><td>${valor}</td></tr>`;
  }).join("");
  return `<table class="diag-table"><tbody>${linhas}</tbody></table>`;
}

function escapeHtml(texto) {
  return String(texto ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}


async function carregarCatalogoPdfsBancoHoras() {
  const alvo = document.getElementById("catalogoPdfsBancoHoras");
  if (!alvo) return;
  const nomeSelecionado = qs("nome") ? qs("nome").value : "";
  const dataBase = alvo.dataset.dataBase || "";
  const matriculaBase = alvo.dataset.matriculaBase || "";
  alvo.innerHTML = `<p>Carregando catálogo dos PDFs lidos...</p>`;
  try {
    const p = new URLSearchParams();
    if (nomeSelecionado) p.set("nome", nomeSelecionado);
    if (dataBase) p.set("data", dataBase);
    if (matriculaBase) p.set("matricula", matriculaBase);
    const resp = await fetch(`/api/catalogo-pdfs-banco-horas?${p.toString()}`);
    const json = await resp.json();
    if (!resp.ok || !json.ok) {
      alvo.innerHTML = `<p class="erro">${escapeHtml(json.erro || `Erro HTTP ${resp.status}`)}</p>`;
      return;
    }
    const d = json.dados || {};
    const linhas = (d.linhas || []).map(l => `
      <tr>
        <td>${escapeHtml(l.arquivo_pdf || "")}</td>
        <td>${escapeHtml(l.nome_pdf || "")}</td>
        <td>${escapeHtml(l.matricula_pdf || "")}</td>
        <td>${escapeHtml(l.matricula_norm_pdf || "")}</td>
        <td>${escapeHtml(l.periodo || "")}</td>
        <td>${escapeHtml(l.emissao || "")}</td>
        <td>${escapeHtml(l.datas_reais_pdf ? "sim" : "não")}</td>
        <td>${escapeHtml(String(l.qtd_datas_lancadas ?? ""))}</td>
        <td>${escapeHtml(l.primeira_data_lancada || "")}</td>
        <td>${escapeHtml(l.ultima_data_lancada || "")}</td>
        <td>${escapeHtml(l.saldo_anterior_pdf || "")}</td>
        <td>${escapeHtml(l.debito_bh_pdf || "")}</td>
        <td>${escapeHtml(l.credito_bh_pdf || "")}</td>
        <td>${escapeHtml(l.saldo_atual_pdf || "")}</td>
        <td>${escapeHtml(l.match_nome ? "sim" : "não")}</td>
        <td>${escapeHtml(l.match_matricula ? "sim" : "não")}</td>
        <td>${escapeHtml(l.dentro_periodo_data ? "sim" : "não")}</td>
        <td>${escapeHtml(l.contem_data_real ? "sim" : "não")}</td>
      </tr>
    `).join("");

    alvo.innerHTML = `
      <div class="diag-catalogo-resumo">
        <strong>PDFs encontrados:</strong> ${escapeHtml(String(d.pdfs_encontrados ?? 0))} &nbsp;|&nbsp;
        <strong>Quadros BH lidos:</strong> ${escapeHtml(String(d.quadros_lidos ?? 0))} &nbsp;|&nbsp;
        <strong>Filtro nome:</strong> ${escapeHtml(d.nome_filtro || "sem filtro")} &nbsp;|&nbsp;
        <strong>Data base:</strong> ${escapeHtml(d.data_filtro || "sem filtro")} &nbsp;|&nbsp;
        <strong>Matrícula:</strong> ${escapeHtml(d.matricula_filtro || "sem filtro")}
      </div>
      <div class="diag-scroll">
        <table class="diag-table diag-table-wide">
          <thead>
            <tr>
              <th>PDF</th><th>Nome no PDF</th><th>Matrícula PDF</th><th>Matrícula norm.</th>
              <th>Período</th><th>Emissão</th><th>Datas reais?</th><th>Qtd datas</th>
              <th>1ª data</th><th>Última data</th><th>Saldo Anterior</th><th>Débito</th>
              <th>Crédito</th><th>Saldo Atual</th><th>Match nome?</th><th>Match matrícula?</th>
              <th>Dentro período?</th><th>Contém data real?</th>
            </tr>
          </thead>
          <tbody>${linhas || `<tr><td colspan="18">Nenhum quadro de Banco de Horas catalogado.</td></tr>`}</tbody>
        </table>
      </div>
    `;
  } catch (e) {
    alvo.innerHTML = `<p class="erro">Erro ao carregar catálogo dos PDFs: ${escapeHtml(e.message)}</p>`;
  }
}

async function carregarDiagnosticoBancoHoras() {
  const paginaEsperada = "extrato_banco_horas";
  const tokenEsperado = tokenAtualPagina();
  const painel = garantirPainelDiagnosticoBancoHoras();
  const nomeSelecionado = qs("nome") ? qs("nome").value : "";
  if (!nomeSelecionado) {
    painel.classList.add("hidden");
    painel.innerHTML = "";
    return;
  }
  painel.classList.remove("hidden");
  painel.innerHTML = `<h3>Diagnóstico cirúrgico do Banco de Horas</h3><p>Carregando diagnóstico...</p>`;
  try {
    const p = new URLSearchParams();
    p.set("nome", nomeSelecionado);
    const resp = await fetch(`/api/diagnostico-banco-horas?${p.toString()}`);
    const json = await resp.json();
    if (!requisicaoAindaValida(paginaEsperada, tokenEsperado)) return;
    if (!resp.ok || !json.ok) {
      painel.innerHTML = `<h3>Diagnóstico cirúrgico do Banco de Horas</h3><p class="erro">${escapeHtml(json.erro || `Erro HTTP ${resp.status}`)}</p>`;
      return;
    }
    const d = json.dados || {};
    const pdf = d.diagnostico_pdf || {};
    const pdfBase = d.pdf_base_ciclo_anterior || {};
    const pdfAtual = d.pdf_fotografia_ciclo_atual || {};
    const candidatos = (pdf.candidatos_mesmo_nome_periodo || []).map(c => `
      <tr>
        <td>${escapeHtml(c.arquivo_pdf || "")}</td>
        <td>${escapeHtml(c.periodo || "")}</td>
        <td>${escapeHtml(c.emissao || "")}</td>
        <td>${escapeHtml(c.datas_reais_pdf ? "sim" : "não")}</td>
        <td>${escapeHtml(c.contem_data_real ? "sim" : "não")}</td>
        <td>${escapeHtml(c.data_max_lancada || "")}</td>
        <td>${escapeHtml(c.saldo_anterior_pdf || "")}</td>
        <td>${escapeHtml(c.saldo_atual_pdf || "")}</td>
      </tr>
    `).join("");

    painel.innerHTML = `
      <h3>Diagnóstico cirúrgico do Banco de Horas</h3>
      <div class="diag-grid">
        <div><strong>Colaborador</strong><span>${escapeHtml(d.nome || "")}</span></div>
        <div><strong>Data base</strong><span>${escapeHtml(d.data_base || "")}</span></div>
        <div><strong>Ciclo</strong><span>${escapeHtml((d.inicio_ciclo || "") + " a " + (d.fim_ciclo || ""))}</span></div>
        <div><strong>PDF base ciclo anterior</strong><span>${escapeHtml(pdfBase.pdf || "não encontrado")}</span></div>
        <div><strong>Saldo base PDF</strong><span>${escapeHtml(pdfBase.saldo_atual_pdf || "vazio")}</span></div>
        <div><strong>PDF fotografia ciclo atual</strong><span>${escapeHtml(pdfAtual.pdf || pdf.pdf_escolhido || "não encontrado")}</span></div>
        <div><strong>Saldo informado PDF</strong><span>${escapeHtml(pdfAtual.saldo_atual_pdf || pdf.saldo_atual_pdf || "vazio")}</span></div>
        <div><strong>Matrícula Excel norm.</strong><span>${escapeHtml((d.linha_normalizada || {}).matricula_normalizada_excel || d.matricula_base || "vazio")}</span></div>
        <div><strong>Origem normalizada</strong><span>${escapeHtml((d.linha_normalizada || {}).origem_saldo_importado || "")}</span></div>
      </div>
      <p class="diag-note">${escapeHtml(d.leitura || "")}</p>
      <details open><summary>Linha bruta do Excel - colunas de saldo</summary>${renderObjetoDiagnostico(d.linha_bruta_saldos || {})}</details>
      <details open><summary>Linha normalizada usada pelo Extrato</summary>${renderObjetoDiagnostico(d.linha_normalizada || {})}</details>
      <details open><summary>PDF base do ciclo anterior usado no Saldo Calculado</summary>${renderObjetoDiagnostico(pdfBase)}</details>
      <details open><summary>PDF/fotografia do ciclo atual usado no Saldo Informado</summary>${renderObjetoDiagnostico(pdfAtual)}</details>
      <details open><summary>Diagnóstico da leitura dos PDFs</summary>${renderObjetoDiagnostico({
        pdfplumber_instalado: pdf.pdfplumber_instalado,
        pdfs_encontrados: pdf.pdfs_encontrados,
        quadros_lidos: pdf.quadros_lidos,
        match_encontrado: pdf.match_encontrado,
        pdf_escolhido: pdf.pdf_escolhido,
        saldo_anterior_pdf: pdf.saldo_anterior_pdf,
        debito_bh_pdf: pdf.debito_bh_pdf,
        credito_bh_pdf: pdf.credito_bh_pdf,
        saldo_atual_pdf: pdf.saldo_atual_pdf,
        motivo: pdf.motivo
      })}</details>
      <details><summary>Candidatos por nome/período encontrados nos PDFs</summary>
        <table class="diag-table">
          <thead><tr><th>PDF</th><th>Período</th><th>Emissão</th><th>Datas reais?</th><th>Contém data?</th><th>Última data</th><th>Saldo Anterior</th><th>Saldo Atual</th></tr></thead>
          <tbody>${candidatos || `<tr><td colspan="8">Nenhum candidato encontrado.</td></tr>`}</tbody>
        </table>
      </details>
      <details open><summary>Catálogo dos PDFs do Banco de Horas - v8.10.18</summary>
        <p class="diag-note">Use este catálogo para conferir como o Robô está lendo nome, matrícula, período, datas reais e saldos de cada PDF. A tabela vem ordenada pelos melhores candidatos ao colaborador selecionado.</p>
        <button class="btn secondary" type="button" onclick="carregarCatalogoPdfsBancoHoras()">Carregar catálogo dos PDFs</button>
        <div id="catalogoPdfsBancoHoras" class="diag-catalogo-placeholder" data-data-base="${escapeHtml(d.data_base || '')}" data-matricula-base="${escapeHtml(d.matricula_base || '')}"><em>Catálogo ainda não carregado.</em></div>
      </details>
    `;
  } catch (e) {
    painel.innerHTML = `<h3>Diagnóstico cirúrgico do Banco de Horas</h3><p class="erro">Erro ao carregar diagnóstico: ${escapeHtml(e.message)}</p>`;
  }
}




async function carregarOpcoesExtratoBancoHoras() {
  const ccEl = qs("cc");
  const nomeEl = qs("nome");
  const turnoEl = qs("turno");
  const dataEl = qs("data");

  const valorNomeAtual = nomeEl ? nomeEl.value : "";
  const valorTurnoAtual = turnoEl ? turnoEl.value : "";
  const valorCcAtual = ccEl ? ccEl.value : "";

  const p = new URLSearchParams();
  if (valorCcAtual) p.set("cc", valorCcAtual);
  if (valorTurnoAtual) p.set("turno", valorTurnoAtual);
  if (typeof datasFiltro !== "undefined" && datasFiltro.length > 0) {
    p.set("datas", datasFiltro.join(","));
  } else if (dataEl && dataEl.value) {
    p.set("data", dataEl.value);
  }

  try {
    const resp = await fetch(`/api/opcoes-extrato-banco-horas?${p.toString()}`);
    if (resp.status === 401) {
      window.location.href = "/login";
      return;
    }
    const json = await resp.json();
    if (!resp.ok || json.erro) {
      console.warn("Falha ao carregar opções do Extrato Banco de Horas", json.erro || resp.status);
      return;
    }

    const nomes = json.nomes || [];
    const turnos = json.turnos || [];
    const ccs = json.ccs || [];

    if (ccEl && ccs.length && ccEl.options.length <= 1) {
      preencherSelect("cc", ccs, "Todos os CCs");
      ccEl.value = valorCcAtual || "";
    }

    preencherSelect("nome", nomes, "Todos os colaboradores");
    if (valorNomeAtual && nomes.includes(valorNomeAtual)) {
      qs("nome").value = valorNomeAtual;
    }

    preencherSelect("turno", turnos, "Todos os turnos");
    if (valorTurnoAtual && turnos.includes(valorTurnoAtual)) {
      qs("turno").value = valorTurnoAtual;
    }

    if (qs("statusCarga") && paginaAtual === "extrato_banco_horas" && !valorNomeAtual) {
      qs("statusCarga").textContent = `${nomes.length} colaborador(es) disponível(is) para filtro.`;
    }
  } catch (e) {
    console.warn("Erro ao carregar opções do Extrato Banco de Horas", e);
  }
}


function prepararDiagnosticoBancoHorasOpcional() {
  // v8.15.17: diagnóstico cirúrgico removido das guias operacionais.
  const painel = document.getElementById("painelDiagnosticoBancoHoras");
  if (painel) {
    painel.innerHTML = "";
    painel.classList.add("hidden");
  }
}


async function carregarExtratoBancoHoras() {
  const paginaEsperada = "extrato_banco_horas";
  const tokenEsperado = tokenAtualPagina();
  paginaAtual = "extrato_banco_horas";
  limparResiduosBancoHoras2ForaDaGuiaBanco();
  dadosConsolidadoAtual = [];

  if (qs("tituloPagina")) qs("tituloPagina").textContent = "Extrato de Banco de Horas";
  if (qs("subtituloPagina")) qs("subtituloPagina").textContent = "Auditoria diária do saldo: saldo base, créditos, débitos e saldo calculado.";
  if (qs("tituloTabela")) qs("tituloTabela").textContent = "Extrato de Banco de Horas";
  if (qs("tbody")) qs("tbody").innerHTML = "";

  ajustarFiltrosExtratoBancoHoras();
  setCabecalhoExtratoBancoHoras();

  if (qs("nome") && qs("nome").options.length <= 1) {
    await carregarOpcoesExtratoBancoHoras();
  }

  const nomeSelecionado = qs("nome") ? qs("nome").value : "";
  // v8.11.27: não bloqueia somente pela ausência de nome. O backend agora
  // consegue aplicar CC/turno/data e autoabrir quando o filtro resultar em um único colaborador.

  const p = new URLSearchParams();
  const ccSelecionado = qs("cc") ? qs("cc").value : "";
  const turnoSelecionado = qs("turno") ? qs("turno").value : "";
  const dataEl = qs("data");

  if (ccSelecionado) p.set("cc", ccSelecionado);
  if (nomeSelecionado) p.set("nome", nomeSelecionado);
  if (turnoSelecionado) p.set("turno", turnoSelecionado);
  if (typeof datasFiltro !== "undefined" && datasFiltro.length > 0) {
    p.set("datas", datasFiltro.join(","));
  } else if (dataEl && dataEl.value) {
    p.set("data", dataEl.value);
  }
  // v8.11.30: evita renderização gigante no navegador.
  p.set("limit", "500");
  p.set("max_colaboradores", "20");

  try {
    if (qs("statusCarga")) qs("statusCarga").textContent = `Carregando Extrato de Banco de Horas${p.toString() ? " com filtros aplicados..." : "..."}`;
    if (controllerExtratoBancoHoras) {
      try { controllerExtratoBancoHoras.abort(); } catch (e) {}
    }
    controllerExtratoBancoHoras = new AbortController();
    const resp = await fetch(`/api/extrato-banco-horas?${p.toString()}`, { signal: controllerExtratoBancoHoras.signal });
    const json = await resp.json();
    if (!requisicaoAindaValida(paginaEsperada, tokenEsperado)) return;

    if (!resp.ok || json.erro) {
      renderCardsExtratoBancoHoras({});
      if (qs("statusCarga")) qs("statusCarga").textContent = json.erro || `Erro HTTP ${resp.status}`;
      if (qs("tbody")) qs("tbody").innerHTML = "";
      return;
    }

    renderCardsExtratoBancoHoras(json.resumo || {});
    renderExtratoBancoHoras(json.dados || []);
    const totalLinhas = Number(json.total_linhas || (json.dados || []).length || 0);
    const qtdLinhas = (json.dados || []).length;
    const aviso = json.aviso ? String(json.aviso) : "";
    const sufixoCorte = json.truncado ? ` Exibindo ${qtdLinhas} de ${totalLinhas} linha(s) para preservar a navegação.` : "";
    if (qs("statusCarga")) qs("statusCarga").textContent = aviso || `${qtdLinhas} linha(s) de extrato carregada(s).${sufixoCorte}`;
  } catch (e) {
    if (e && e.name === "AbortError") return;
    renderCardsExtratoBancoHoras({});
    if (qs("statusCarga")) qs("statusCarga").textContent = "Erro ao carregar extrato: " + e.message;
    if (qs("tbody")) qs("tbody").innerHTML = "";
  }
}




function setCabecalhoAuditoriaPremium() {
  if (qs("tituloPagina")) qs("tituloPagina").textContent = "Auditoria";
  if (qs("subtituloPagina")) qs("subtituloPagina").textContent = "Rastreabilidade de fechamentos.";
  setVisibilidadeDashboardExecutivoPremium(true);
}

function paramsAuditoriaPremium() {
  const p = new URLSearchParams();
  const cc = qs("auditCc") ? qs("auditCc").value : "";
  const nome = qs("auditNome") ? qs("auditNome").value : "";
  const turno = qs("auditTurno") ? qs("auditTurno").value : "";
  const data = qs("auditData") ? qs("auditData").value : "";
  const emissao = qs("auditEmissao") ? qs("auditEmissao").value : "";
  if (cc) p.set("cc", cc);
  if (nome) p.set("nome", nome);
  if (turno) p.set("turno", turno);
  if (data) p.set("data", data);
  if (emissao) p.set("emissao", emissao);
  return p;
}

function renderAuditoriaPremium(payload) {
  const painel = qs("dashboardExecutivoPremium");
  if (!painel) return;
  const dados = payload.dados || [];
  const k = payload.kpis || {};
  const op = payload.opcoes || {};
  const opt = (lista, vazio) => `<option value="">${vazio}</option>` + (lista || []).map(v => `<option value="${v}">${v}</option>`).join("");
  const filtrosExistentes = {
    cc: qs("auditCc") ? qs("auditCc").value : "",
    nome: qs("auditNome") ? qs("auditNome").value : "",
    turno: qs("auditTurno") ? qs("auditTurno").value : "",
    data: qs("auditData") ? qs("auditData").value : "",
    emissao: qs("auditEmissao") ? qs("auditEmissao").value : "",
  };
  painel.innerHTML = `
    <div class="exec-hero audit-hero-compact">
      <div>
        <h3>Rastreabilidade de Fechamentos</h3>
      </div>
      <div class="exec-periodo">
        <strong>${k.linhas || 0} linha(s)</strong>
        <small>${k.colaboradores || 0} colaborador(es) · ${k.emissoes || 0} emissão(ões)</small>
      </div>
    </div>
    <section class="filters audit-filters">
      <select id="auditCc">${opt(op.ccs, "Todos os CCs")}</select>
      <select id="auditNome">${opt(op.nomes, "Todos os colaboradores")}</select>
      <select id="auditTurno">${opt(op.turnos, "Todos os turnos")}</select>
      <input id="auditData" type="date" />
      <input id="auditEmissao" type="date" title="Data Emissão PDF" />
      <button class="btn secondary" onclick="limparAuditoriaPremium()">Limpar</button>
      <button class="btn" onclick="carregarAuditoriaPremium()">Aplicar</button>
    </section>
    ${payload.erro ? `<div class="exec-alert danger"><strong>Auditoria indisponível</strong><span>${escapeHtml(payload.erro)}</span></div>` : ""}
    ${payload.aviso ? `<div class="exec-alert warning"><strong>Modo compatível</strong><span>${escapeHtml(payload.aviso)}</span></div>` : ""}
    <div class="exec-kpi-grid">
      ${cardResumoExecutivo("Linhas históricas", k.linhas || 0, k.limitado ? `exibindo ${k.limite}` : "sem corte", "")}
      ${cardResumoExecutivo("Colaboradores", k.colaboradores || 0, "no filtro atual", "")}
      ${cardResumoExecutivo("Datas", k.datas || 0, "dias auditados", "")}
      ${cardResumoExecutivo("Emissões", k.emissoes || 0, "fotografias PDF", "warning")}
    </div>
    <section class="exec-panel">
      <div class="exec-panel-title"><div><h4>Linhas de Auditoria</h4><p>Mostra até ${k.limite || 600} linhas para preservar a navegação.</p></div></div>
      <div class="table-wrap"><table class="dashboard-table"><thead><tr>
        <th>Data</th><th>Data Emissão PDF</th><th>Colaborador</th><th>CC</th><th>Turno</th><th>H.E.</th><th>Absent.</th><th>Saldo Calculado</th><th>Saldo Informado PDF</th><th>Saldo Atual</th><th>Observação</th><th>Arquivo Origem</th>
      </tr></thead><tbody>
        ${dados.length ? dados.map(r => `<tr><td>${r.data || ""}</td><td>${r.data_emissao_pdf || ""}</td><td>${r.nome || ""}</td><td>${r.cc || ""}</td><td>${r.turno || ""}</td><td>${r.he || ""}</td><td>${r.absent || ""}</td><td>${r.saldo_calculado || ""}</td><td>${r.saldo_informado_pdf || ""}</td><td>${r.saldo_atual || ""}</td><td>${r.observacao || ""}</td><td>${r.arquivo_origem || ""}</td></tr>`).join("") : `<tr><td colspan="12">Nenhuma linha encontrada para os filtros.</td></tr>`}
      </tbody></table></div>
    </section>
  `;
  if (qs("auditCc")) qs("auditCc").value = filtrosExistentes.cc;
  if (qs("auditNome")) qs("auditNome").value = filtrosExistentes.nome;
  if (qs("auditTurno")) qs("auditTurno").value = filtrosExistentes.turno;
  if (qs("auditData")) qs("auditData").value = filtrosExistentes.data;
  if (qs("auditEmissao")) qs("auditEmissao").value = filtrosExistentes.emissao;
}

async function carregarAuditoriaPremium() {
  const paginaEsperada = "auditoria_premium";
  const tokenEsperado = tokenAtualPagina();
  paginaAtual = "auditoria_premium";
  marcarPaginaAtualClasse();
  setCabecalhoAuditoriaPremium();
  if (qs("cardsContainer")) qs("cardsContainer").classList.add("hidden");
  if (qs("filtros")) qs("filtros").classList.add("hidden");
  if (qs("filtrosDashboard")) qs("filtrosDashboard").classList.add("hidden");
  if (qs("tabelaConsolidado")) qs("tabelaConsolidado").classList.add("hidden");
  if (qs("tabelaDashboard")) qs("tabelaDashboard").classList.add("hidden");
  if (qs("controleOrdenacaoBancoHoras")) qs("controleOrdenacaoBancoHoras").classList.add("hidden");
  if (qs("tituloTabela")) qs("tituloTabela").textContent = "Auditoria";
  if (qs("statusCarga")) qs("statusCarga").textContent = "Carregando auditoria...";
  const { controller, timer } = criarControllerNavegacao(12000);
  try {
    const resp = await fetch(`/api/auditoria-premium?${paramsAuditoriaPremium().toString()}`, { signal: controller.signal });
    const json = await resp.json();
    if (!requisicaoAindaValida(paginaEsperada, tokenEsperado)) return;
    renderAuditoriaPremium(json);
    if (qs("statusCarga")) qs("statusCarga").textContent = json.erro || `${(json.kpis || {}).linhas || 0} linha(s) auditável(is).`;
  } catch (e) {
    if (!requisicaoAindaValida(paginaEsperada, tokenEsperado)) return;
    renderAuditoriaPremium({erro: "Erro ao carregar auditoria: " + e.message, dados: [], kpis: {}, opcoes: {}});
    if (qs("statusCarga")) qs("statusCarga").textContent = e.name === "AbortError" ? "Carregamento da Auditoria cancelado pela troca de guia." : "Erro ao carregar Auditoria.";
  } finally {
    finalizarControllerNavegacao(controller, timer);
  }
}

function limparAuditoriaPremium() {
  ["auditCc", "auditNome", "auditTurno", "auditData", "auditEmissao"].forEach(id => { if (qs(id)) qs(id).value = ""; });
  carregarAuditoriaPremium();
}

function abrirAuditoriaPremium(el) {
  iniciarTransicaoPagina("auditoria_premium");
  paginaAtual = "auditoria_premium";
  marcarPaginaAtualClasse();
  setVisibilidadeDashboardExecutivoPremium(true);
  dadosConsolidadoAtual = [];
  document.querySelectorAll(".menu button").forEach(b => b.classList.remove("active"));
  if (el) el.classList.add("active");
  carregarAuditoriaPremium();
}

function abrirDashboardExecutivo(el) {
  iniciarTransicaoPagina("dashboard_executivo");
  paginaAtual = "dashboard_executivo";
  marcarPaginaAtualClasse();
  setVisibilidadeDashboardExecutivoPremium(true);
  dadosConsolidadoAtual = [];

  document.querySelectorAll(".menu button").forEach(b => b.classList.remove("active"));
  if (el) el.classList.add("active");

  if (qs("filtros")) qs("filtros").classList.remove("hidden");
  if (qs("filtrosDashboard")) qs("filtrosDashboard").classList.add("hidden");
  if (qs("boxFiltroTipo")) qs("boxFiltroTipo").classList.add("hidden");
  if (qs("controleOrdenacaoBancoHoras")) qs("controleOrdenacaoBancoHoras").classList.add("hidden");
  if (qs("tabelaDashboard")) qs("tabelaDashboard").classList.add("hidden");
  if (qs("tabelaConsolidado")) qs("tabelaConsolidado").classList.remove("hidden");
  ocultarFiltroDataPrincipalPremium();

  carregarDashboardExecutivo();
}

function setCabecalhoDashboardExecutivo() {
  // O Dashboard Executivo usa cards, rankings e barras visuais.
  // A tabela tradicional fica preservada para as demais guias.
  setVisibilidadeDashboardExecutivoPremium(true);
}

function valorSeguro(v, padrao = "") {
  return (v === null || v === undefined) ? padrao : v;
}

function classeImpactoDashboard(valorMinutos, tipo = "neutro") {
  const v = Number(valorMinutos || 0);
  if (tipo === "negativo") return v < 0 ? "danger" : "success";
  if (v > 0 && tipo === "alerta") return "warning";
  return "neutral";
}

function itemRankingExecutivo(r, maxAbs, campoValor = "valor") {
  const bruto = Math.abs(Number(r.minutos ?? r.valor_min ?? 0));
  const pct = maxAbs > 0 ? Math.max(4, Math.min(100, (bruto / maxAbs) * 100)) : 0;
  return `
    <div class="exec-rank-row">
      <div class="exec-rank-main">
        <strong>${valorSeguro(r.nome, "-")}</strong>
        <span>${valorSeguro(r.cc, "")} ${r.funcao ? "• " + r.funcao : ""} ${r.turno ? "• " + r.turno : ""}</span>
      </div>
      <div class="exec-rank-value">${valorSeguro(r[campoValor], "00:00")}</div>
      <div class="exec-bar"><i style="width:${pct.toFixed(1)}%"></i></div>
    </div>
  `;
}

function blocoRankingExecutivo(titulo, subtitulo, linhas, vazio) {
  linhas = linhas || [];
  const maxAbs = Math.max(...linhas.map(r => Math.abs(Number(r.minutos ?? r.valor_min ?? 0))), 0);
  return `
    <section class="exec-panel">
      <div class="exec-panel-title">
        <div>
          <h4>${titulo}</h4>
          <p>${subtitulo}</p>
        </div>
      </div>
      <div class="exec-rank-list">
        ${linhas.length ? linhas.map(r => itemRankingExecutivo(r, maxAbs)).join("") : `<div class="exec-empty">${vazio}</div>`}
      </div>
    </section>
  `;
}

function cardResumoExecutivo(label, valor, detalhe, classe = "") {
  return `
    <div class="exec-kpi ${classe}">
      <span>${label}</span>
      <strong>${valor}</strong>
      <small>${detalhe || ""}</small>
    </div>
  `;
}

function renderResumoPorGrupo(titulo, subtitulo, linhas, tipo) {
  linhas = linhas || [];
  const max = Math.max(...linhas.map(r => Math.abs(Number(r.saldo_min ?? r.he_min ?? r.absent_min ?? 0))), 0);
  return `
    <section class="exec-panel">
      <div class="exec-panel-title">
        <div>
          <h4>${titulo}</h4>
          <p>${subtitulo}</p>
        </div>
      </div>
      <div class="exec-group-list">
        ${linhas.length ? linhas.map(r => {
          const base = tipo === "cc" ? Math.abs(Number(r.saldo_min || 0)) : Math.abs(Number(r.he_min || 0)) + Math.abs(Number(r.absent_min || 0));
          const pct = max > 0 ? Math.max(4, Math.min(100, (base / max) * 100)) : 0;
          return `
            <div class="exec-group-row">
              <div>
                <strong>${valorSeguro(r.rotulo || r.grupo, "-")}</strong>
                <span>${valorSeguro(r.colaboradores, 0)} colaborador(es) • ${valorSeguro(r.registros, 0)} registro(s)</span>
              </div>
              <div class="exec-group-metrics">
                <b>Saldo ${valorSeguro(r.saldo, "00,00")}</b>
                <span>HE ${valorSeguro(r.he, "00:00")} • Aus. ${valorSeguro(r.absent, "00:00")}</span>
              </div>
              <div class="exec-bar"><i style="width:${pct.toFixed(1)}%"></i></div>
            </div>
          `;
        }).join("") : `<div class="exec-empty">Sem dados para este filtro.</div>`}
      </div>
    </section>
  `;
}

function renderAlertasExecutivos(alertas) {
  alertas = alertas || [];
  return `
    <section class="exec-panel exec-alert-panel">
      <div class="exec-panel-title">
        <div>
          <h4>Alertas Gerenciais</h4>
          <p>Sinais que merecem tratativa antes de virar escalonamento.</p>
        </div>
      </div>
      <div class="exec-alert-list">
        ${alertas.length ? alertas.map(a => `
          <div class="exec-alert ${a.nivel || "info"}">
            <strong>${a.titulo}</strong>
            <span>${a.descricao}</span>
          </div>
        `).join("") : `<div class="exec-alert ok"><strong>Sem alertas críticos</strong><span>Os indicadores filtrados não apontam concentração relevante de risco.</span></div>`}
      </div>
    </section>
  `;
}


function renderTendenciasOperacionais(tendencias) {
  tendencias = tendencias || {};
  const comps = tendencias.comparativos || [];
  const atual = tendencias.atual || {};
  const anterior = tendencias.anterior || {};
  return `
    <section class="exec-panel">
      <div class="exec-panel-title">
        <div>
          <h4>Tendência Operacional</h4>
          <p>Comparação entre a metade inicial e a metade final do período filtrado.</p>
        </div>
        <span class="exec-chip">${anterior.inicio || "-"} a ${anterior.fim || "-"} → ${atual.inicio || "-"} a ${atual.fim || "-"}</span>
      </div>
      <div class="exec-list">
        ${comps.length ? comps.map(c => `
          <div class="exec-row">
            <span><strong>${c.indicador}</strong><small>Anterior: ${c.anterior} · Atual: ${c.atual}</small></span>
            <b class="${c.direcao === "alta" ? "negative" : c.direcao === "queda" ? "positive" : ""}">${c.variacao}</b>
          </div>
        `).join("") : `<div class="exec-empty">Período insuficiente para comparar tendência.</div>`}
      </div>
    </section>
  `;
}

function renderAlertasInteligentes(alertas) {
  alertas = alertas || [];
  const label = {critico: "Crítico", alto: "Alto", medio: "Médio", baixo: "Baixo"};
  return `
    <section class="exec-panel exec-alert-panel">
      <div class="exec-panel-title">
        <div>
          <h4>Central de Alertas Inteligentes</h4>
          <p>Classificação automática de risco por colaborador.</p>
        </div>
      </div>
      <div class="exec-alert-list">
        ${alertas.length ? alertas.map(a => `
          <div class="exec-alert ${a.nivel === "critico" ? "danger" : a.nivel === "alto" ? "warning" : a.nivel === "medio" ? "info" : "ok"}">
            <strong>${label[a.nivel] || a.nivel} · ${a.nome}</strong>
            <span>${a.motivo} CC ${a.cc || "-"} · Turno ${a.turno || "-"} · Banco ${a.saldo || "00,00"} · H.E. ${a.he || "00:00"} · Ausências ${a.ausencias || "00:00"}</span>
          </div>
        `).join("") : `<div class="exec-alert ok"><strong>Sem risco relevante</strong><span>Nenhum colaborador atingiu os critérios de alerta inteligente.</span></div>`}
      </div>
    </section>
  `;
}

function renderDashboardExecutivoPremium(payload) {
  const painel = qs("dashboardExecutivoPremium");
  if (!painel) return;

  const k = payload.kpis || {};
  const rankings = payload.rankings || {};
  const resumos = payload.resumos || {};
  const periodo = payload.periodo || {};
  const registros = k.registros ?? 0;
  const taxaRevisao = k.taxa_revisao ?? "0%";

  painel.innerHTML = `
    <div class="exec-hero">
      <div>
        <h3>Saúde Operacional do Ponto</h3>
        <p>Leitura executiva de banco de horas, horas extras, ausências e registros que exigem atenção.</p>
      </div>
      <div class="exec-periodo">
        <span>Período analisado</span>
        <strong>${periodo.inicio || "-"} ${periodo.fim && periodo.fim !== periodo.inicio ? "a " + periodo.fim : ""}</strong>
        <small>${registros} registro(s) na visão atual</small>
      </div>
    </div>

    <div class="exec-kpi-grid">
      ${cardResumoExecutivo("Colaboradores", k.colaboradores ?? 0, "pessoas avaliadas", "") }
      ${cardResumoExecutivo("Banco Positivo", k.banco_positivo ?? "00,00", `${k.colab_saldo_positivo ?? 0} colaborador(es)`, "positive") }
      ${cardResumoExecutivo("Banco Negativo", k.banco_negativo ?? "00,00", `${k.colab_saldo_negativo ?? 0} colaborador(es)`, "negative") }
      ${cardResumoExecutivo("Horas Extras", k.he_total ?? "00:00", `${k.dias_com_he ?? 0} dia(s) com HE`, "warning") }
      ${cardResumoExecutivo("Ausências", k.ausencia_total ?? "00:00", `${k.dias_com_ausencia ?? 0} dia(s) com ausência`, "danger") }
      ${cardResumoExecutivo("Revisão", taxaRevisao, `${k.registros_revisao ?? 0} registro(s) com alerta`, "") }
    </div>

    <div class="exec-grid exec-grid-3">
      ${blocoRankingExecutivo("Maiores Saldos Positivos", "Quem concentra créditos no Banco de Horas.", rankings.banco_positivo, "Sem saldo positivo no filtro atual.")}
      ${blocoRankingExecutivo("Maiores Saldos Negativos", "Quem exige atenção por déficit acumulado.", rankings.banco_negativo, "Sem saldo negativo no filtro atual.")}
      ${blocoRankingExecutivo("Top Horas Extras", "Maiores volumes de H.E. no período filtrado.", rankings.horas_extras, "Sem horas extras no filtro atual.")}
    </div>

    <div class="exec-grid exec-grid-2">
      ${blocoRankingExecutivo("Top Ausências", "Colaboradores com maior volume de ausência.", rankings.ausencias, "Sem ausências no filtro atual.")}
      ${renderAlertasExecutivos(payload.alertas)}
    </div>

    <div class="exec-grid exec-grid-2">
      ${renderTendenciasOperacionais(payload.tendencias)}
      ${renderAlertasInteligentes(payload.alertas_inteligentes)}
    </div>

    <div class="exec-grid exec-grid-2">
      ${renderResumoPorGrupo("Resumo por Centro de Custo", "Saldo, H.E. e ausências por CC.", resumos.por_cc, "cc")}
      ${renderResumoPorGrupo("Resumo por Turno", "Concentração operacional por turno.", resumos.por_turno, "turno")}
    </div>
  `;
}

function renderCardsDashboardExecutivo(k) {
  const container = qs("cardsContainer");
  if (!container) return;
  k = k || {};
  container.innerHTML = `
    <div class="card"><span>Registros</span><strong>${k.registros ?? 0}</strong></div>
    <div class="card"><span>Colaboradores</span><strong>${k.colaboradores ?? 0}</strong></div>
    <div class="card"><span>Banco Negativo</span><strong>${k.banco_negativo ?? "00,00"}</strong></div>
    <div class="card"><span>Horas Extras</span><strong>${k.he_total ?? "00:00"}</strong></div>
    <div class="card"><span>Ausências</span><strong>${k.ausencia_total ?? "00:00"}</strong></div>
    <div class="card"><span>Revisão</span><strong>${k.taxa_revisao ?? "0%"}</strong></div>
  `;
  container.dataset.layoutAtual = "dashboard_executivo";
}

function renderDashboardExecutivo(linhas) {
  // Mantido por compatibilidade com versões anteriores.
  const painel = qs("dashboardExecutivoPremium");
  if (painel && (!linhas || !linhas.length)) {
    painel.innerHTML = `<div class="exec-empty">Nenhum indicador executivo encontrado para os filtros selecionados.</div>`;
  }
}

async function carregarDashboardExecutivo() {
  const paginaEsperada = "dashboard_executivo";
  const tokenEsperado = tokenAtualPagina();
  paginaAtual = "dashboard_executivo";
  marcarPaginaAtualClasse();
  setVisibilidadeDashboardExecutivoPremium(true);
  dadosConsolidadoAtual = [];

  if (qs("tituloPagina")) qs("tituloPagina").textContent = "Dashboard Executivo";
  if (qs("subtituloPagina")) qs("subtituloPagina").textContent = "Leitura gerencial consolidada do Banco de Horas, horas extras, ausências e riscos operacionais.";
  if (qs("tituloTabela")) qs("tituloTabela").textContent = "Dashboard Executivo";
  if (qs("statusCarga")) qs("statusCarga").textContent = "Carregando Dashboard Executivo...";

  if (typeof restaurarFiltrosPadraoPrincipal === "function") restaurarFiltrosPadraoPrincipal();
  paginaAtual = "dashboard_executivo";
  marcarPaginaAtualClasse();
  setVisibilidadeDashboardExecutivoPremium(true);

  if (qs("filtros")) qs("filtros").classList.remove("hidden");
  if (qs("filtrosDashboard")) qs("filtrosDashboard").classList.add("hidden");
  if (qs("boxFiltroTipo")) qs("boxFiltroTipo").classList.add("hidden");
  if (qs("controleOrdenacaoBancoHoras")) qs("controleOrdenacaoBancoHoras").classList.add("hidden");
  if (qs("tabelaDashboard")) qs("tabelaDashboard").classList.add("hidden");
  if (qs("tabelaConsolidado")) qs("tabelaConsolidado").classList.add("hidden");
  ocultarFiltroDataPrincipalPremium();

  setCabecalhoDashboardExecutivo();

  ocultarFiltroDataPrincipalPremium();

  const p = new URLSearchParams();
  if (qs("cc") && qs("cc").value) p.set("cc", qs("cc").value);
  if (qs("nome") && qs("nome").value) p.set("nome", qs("nome").value);
  if (qs("turno") && qs("turno").value) p.set("turno", qs("turno").value);

  const { controller, timer: timeoutDashboard } = criarControllerNavegacao(12000);

  try {
    const resp = await fetch(`/api/dashboard-executivo?${p.toString()}`, { signal: controller.signal });
    const json = await resp.json();
    if (!requisicaoAindaValida(paginaEsperada, tokenEsperado)) return;

    if (!resp.ok || json.erro) {
      renderCardsDashboardExecutivo({});
      renderDashboardExecutivoPremium({});
      if (qs("statusCarga")) qs("statusCarga").textContent = json.erro || `Erro HTTP ${resp.status}`;
      return;
    }

    // v8.11.20: renderização desacoplada para preservar a navegação.
    requestAnimationFrame(() => {
      if (!requisicaoAindaValida(paginaEsperada, tokenEsperado)) return;
      renderCardsDashboardExecutivo(json.kpis || {});
      renderDashboardExecutivoPremium(json);
      if (qs("statusCarga")) qs("statusCarga").textContent = `${json.kpis?.registros ?? 0} registro(s) analisado(s).`;
    });
  } catch (e) {
    if (!requisicaoAindaValida(paginaEsperada, tokenEsperado)) return;
    renderCardsDashboardExecutivo({});
    renderDashboardExecutivoPremium({});
    if (qs("statusCarga")) {
      qs("statusCarga").textContent = e.name === "AbortError"
        ? "Dashboard demorou demais para responder. A navegação foi preservada; tente atualizar a página ou reduzir o filtro."
        : "Erro ao carregar Dashboard Executivo: " + e.message;
    }
  } finally {
    finalizarControllerNavegacao(controller, timeoutDashboard);
  }
}


// ============================================================
// ALERTAS AUTOMÁTICOS - V8.7.0
// ============================================================
function abrirAlertasAutomaticos(el) {
  iniciarTransicaoPagina("alertas_automaticos");
  paginaAtual = "alertas_automaticos";
  marcarPaginaAtualClasse();
  setVisibilidadeDashboardExecutivoPremium(true);
  dadosConsolidadoAtual = [];

  document.querySelectorAll(".menu button").forEach(b => b.classList.remove("active"));
  if (el) el.classList.add("active");

  if (qs("tituloPagina")) qs("tituloPagina").textContent = "Central de Alertas Gerenciais";
  if (qs("subtituloPagina")) {
    qs("subtituloPagina").textContent = "";
    qs("subtituloPagina").innerHTML = "";
  }
  ativarBlindagemSubtituloCentralAlertasV8214();
  limparSubtituloCentralAlertasV8214();
  if (qs("tituloTabela")) qs("tituloTabela").textContent = "Central de Alertas Gerenciais";
  if (qs("statusCarga")) qs("statusCarga").textContent = "Carregando Central de Alertas Gerenciais...";

  if (qs("filtros")) qs("filtros").classList.remove("hidden");
  if (qs("filtrosDashboard")) qs("filtrosDashboard").classList.add("hidden");
  if (qs("boxFiltroTipo")) qs("boxFiltroTipo").classList.add("hidden");
  if (qs("controleOrdenacaoBancoHoras")) qs("controleOrdenacaoBancoHoras").classList.add("hidden");
  if (qs("tabelaDashboard")) qs("tabelaDashboard").classList.add("hidden");
  if (qs("tabelaConsolidado")) qs("tabelaConsolidado").classList.add("hidden");
  ocultarFiltroDataPrincipalPremium();

  carregarAlertasAutomaticos();
  setTimeout(limparSubtituloCentralAlertasV8214, 50);
  setTimeout(limparSubtituloCentralAlertasV8214, 300);
  setTimeout(limparSubtituloCentralAlertasV8214, 1000);
}

function alertaKpi(label, valor, detalhe, classe = "") {
  return `
    <div class="alert-kpi ${classe}">
      <span>${label}</span>
      <strong>${valor}</strong>
      <small>${detalhe || ""}</small>
    </div>
  `;
}

function renderCategoriasAlertas(categorias) {
  categorias = categorias || [];
  return `
    <section class="exec-panel alertas-panel">
      <div class="exec-panel-title">
        <div>
          <h4>Mapa Gerencial de Riscos</h4>
        </div>
      </div>
      <div class="alert-category-grid">
        ${categorias.length ? categorias.map(c => `
          <div class="alert-category ${c.nivel || "info"}">
            <strong>${c.qtd}</strong>
            <div><b>${c.titulo}</b><span>${c.descricao || ""}</span></div>
          </div>
        `).join("") : `<div class="exec-empty">Nenhuma categoria de risco ativa neste filtro.</div>`}
      </div>
    </section>
  `;
}

function renderColaboradoresAlertas(colaboradores) {
  colaboradores = colaboradores || [];
  return `
    <section class="exec-panel alertas-panel">
      <div class="exec-panel-title">
        <div>
          <h4>Fila de Tratativa Gerencial</h4>
          <p>Colaboradores priorizados por score operacional, com sinais para ação.</p>
        </div>
      </div>
      <div class="alert-person-list">
        ${colaboradores.length ? colaboradores.map(c => `
          <div class="alert-person ${c.classe || "info"}">
            <div class="alert-score"><strong>${c.score}</strong><span>${c.nivel}</span></div>
            <div class="alert-person-main">
              <strong>${c.nome}</strong>
              <span>${c.cc || "-"} • ${c.funcao || "-"} • ${c.turno || "-"}</span>
              <div class="alert-metrics">
                <b>Saldo ${c.saldo}</b><b>HE ${c.he}</b><b>Aus. ${c.ausencia}</b><b>Revisão ${c.qtd_revisao}</b>
              </div>
              <div class="alert-tags">
                ${(c.sinais || []).map(s => `<em class="${s.nivel || "info"}" title="${s.detalhe || ""}">${s.tipo}</em>`).join("")}
              </div>
            </div>
          </div>
        `).join("") : `<div class="exec-empty">Sem colaboradores em fila de risco para os filtros selecionados.</div>`}
      </div>
    </section>
  `;
}

function renderOcorrenciasAlertas(ocorrencias) {
  ocorrencias = ocorrencias || [];
  return `
    <section class="exec-panel alertas-panel">
      <div class="exec-panel-title">
        <div>
          <h4>Ocorrências Recentes</h4>
          <p>Registros pontuais que explicam os alertas.</p>
        </div>
      </div>
      <div class="alert-occ-list">
        ${ocorrencias.length ? ocorrencias.slice(0, 18).map(o => `
          <div class="alert-occ-row">
            <span>${o.data || "-"}</span>
            <strong>${o.nome || "-"}</strong>
            <small>${o.cc || "-"} • ${o.turno || "-"}</small>
            <em>${o.motivo || "-"}</em>
          </div>
        `).join("") : `<div class="exec-empty">Nenhuma ocorrência pontual encontrada.</div>`}
      </div>
    </section>
  `;
}

function renderRecomendacoesAlertas(recomendacoes) {
  recomendacoes = recomendacoes || [];
  return `
    <section class="exec-panel alertas-panel">
      <div class="exec-panel-title">
        <div>
          <h4>Ações Recomendadas</h4>
          <p>Próximos passos sugeridos pelo motor de alertas.</p>
        </div>
      </div>
      <div class="alert-recommend-list">
        ${recomendacoes.map(r => `<div class="alert-recommend">${r}</div>`).join("")}
      </div>
    </section>
  `;
}

async function enviarAlertasJornadaAgora() {
  if (!usuarioAtualEhAdmin()) {
    alert("Acesso negado. O envio manual de alertas exige perfil Administrador.");
    return;
  }
  const status = qs("statusCarga");
  if (status) status.textContent = "Verificando e enviando alertas de jornada...";
  try {
    const resp = await fetch("/api/notificacoes-jornada/enviar-agora", {
      method: "POST",
      headers: { "Accept": "application/json" }
    });

    const contentType = resp.headers.get("content-type") || "";
    let json = null;

    if (contentType.includes("application/json")) {
      json = await resp.json();
    } else {
      const texto = await resp.text();
      const resumoHtml = texto.replace(/\s+/g, " ").slice(0, 500);
      throw new Error(`O servidor devolveu HTML em vez de JSON. Status HTTP ${resp.status}. Detalhe: ${resumoHtml}`);
    }

    if (!resp.ok || json.ok === false) {
      const detalhe = json.erro || json.mensagem || `Status HTTP ${resp.status}`;
      const resumoEnviosErro = (json.envios || []).map(e => `${e.canal}: ${e.mensagem}`).join(" | ");
      throw new Error(`${detalhe}${resumoEnviosErro ? " | " + resumoEnviosErro : ""}`);
    }

    const resumoEnvios = (json.envios || []).map(e => `${e.canal}: ${e.mensagem}`).join(" | ");
    const msg = json.mensagem || "Verificação concluída.";
    if (status) status.textContent = `${msg} Novas ocorrências: ${json.novas ?? 0}. ${resumoEnvios}`;
    alert(`${msg}\n\nOcorrências novas: ${json.novas ?? 0}\nDetectadas na base: ${json.total_detectadas ?? 0}\n${resumoEnvios}`);
  } catch (e) {
    const msg = "Erro ao enviar alertas de jornada: " + e.message;
    if (status) status.textContent = msg;
    alert(msg);
  }
}

function renderPainelNotificacoesJornada(status) {
  status = status || {};
  const emailOn = !!status.email_ativo;
  const whatsOn = !!status.whatsapp_ativo;
  const autoOn = !!status.enviar_ao_importar_excel;
  const ultimo = status.ultimo_historico;
  const ultimoTexto = ultimo ? `Último envio: ${ultimo.quando} — ${ultimo.quantidade} ocorrência(s).` : "Nenhum envio registrado ainda.";
  return `
    <div class="alert-notification-card">
      <div>
        <h4>Notificações automáticas de jornada</h4>
        <p>Disparo ao importar o Excel quando houver nova Interjornada ou Violação de Jornada. ${ultimoTexto}</p>
      </div>
      <div class="alert-notification-actions">
        <span class="alert-channel-pill ${emailOn ? "on" : "off"}">E-mail ${emailOn ? "ativo" : "inativo"}</span>
        <span class="alert-channel-pill ${whatsOn ? "on" : "off"}">WhatsApp ${whatsOn ? "ativo" : "inativo"}</span>
        <span class="alert-channel-pill ${autoOn ? "on" : "off"}">Ao importar ${autoOn ? "ativo" : "inativo"}</span>
        <button class="alert-notification-button requires-admin" onclick="abrirConfigNotificacoesJornada()">Configurar notificações</button><button class="alert-notification-button requires-admin" onclick="enviarAlertasJornadaAgora()">Enviar alertas agora</button>
      </div>
    </div>
  `;
}

function renderAlertasAutomaticos(payload) {
  const painel = qs("dashboardExecutivoPremium");
  if (!painel) return;
  const k = payload.kpis || {};
  painel.innerHTML = `
    <div class="alert-hero">
      <div>
        <h3>Central de Alertas</h3>
      </div>
      <div class="alert-seal">
        <span>Score médio</span>
        <strong>${k.score_medio ?? 0}</strong>
        <small>0 a 100</small>
      </div>
    </div>

    ${renderPainelNotificacoesJornada(payload.notificacoes_jornada)}

    <div class="alert-kpi-grid">
      ${alertaKpi("Críticos", k.alertas_criticos ?? 0, "prioridade máxima", "danger")}
      ${alertaKpi("Em atenção", k.alertas_atencao ?? 0, "acompanhar", "warning")}
      ${alertaKpi("Colaboradores", k.colaboradores_risco ?? 0, "com algum sinal", "info")}
      ${alertaKpi("Ocorrências", k.ocorrencias ?? 0, "registros explicativos", "")}
    </div>

    <div class="exec-grid exec-grid-2">
      ${renderCategoriasAlertas(payload.categorias)}
      ${renderRecomendacoesAlertas(payload.recomendacoes)}
    </div>

    <div class="exec-grid exec-grid-2 alert-main-grid">
      ${renderColaboradoresAlertas(payload.colaboradores)}
      ${renderOcorrenciasAlertas(payload.ocorrencias)}
    </div>
  `;
}

async function carregarAlertasAutomaticos() {
  const paginaEsperada = "alertas_automaticos";
  const tokenEsperado = tokenAtualPagina();
  paginaAtual = "alertas_automaticos";
  marcarPaginaAtualClasse();
  setVisibilidadeDashboardExecutivoPremium(true);

  ocultarFiltroDataPrincipalPremium();

  const p = new URLSearchParams();
  if (qs("cc") && qs("cc").value) p.set("cc", qs("cc").value);
  if (qs("nome") && qs("nome").value) p.set("nome", qs("nome").value);
  if (qs("turno") && qs("turno").value) p.set("turno", qs("turno").value);

  try {
    const resp = await fetch(`/api/alertas-automaticos?${p.toString()}`);
    const json = await resp.json();
    if (!requisicaoAindaValida(paginaEsperada, tokenEsperado)) return;
    json.notificacoes_jornada = { email_ativo: false, whatsapp_ativo: false, enviar_ao_importar_excel: false };
    if (!resp.ok || json.erro) {
      renderCardsDashboardExecutivo({});
      renderAlertasAutomaticos({});
      if (qs("statusCarga")) qs("statusCarga").textContent = json.erro || `Erro HTTP ${resp.status}`;
      return;
    }
    renderCardsDashboardExecutivo({
      registros: json.kpis?.ocorrencias ?? 0,
      colaboradores: json.kpis?.colaboradores_risco ?? 0,
      banco_negativo: json.kpis?.alertas_criticos ?? 0,
      he_total: json.kpis?.alertas_atencao ?? 0,
      ausencia_total: json.kpis?.score_medio ?? 0,
      taxa_revisao: `${json.kpis?.ocorrencias ?? 0}`,
    });
    renderAlertasAutomaticos(json);
    if (qs("subtituloPagina")) {
      qs("subtituloPagina").textContent = "";
      qs("subtituloPagina").innerHTML = "";
    }
    limparSubtituloCentralAlertasV8214();
    if (qs("statusCarga")) qs("statusCarga").textContent = `${json.kpis?.colaboradores_risco ?? 0} colaborador(es) com sinal de alerta.`;

    // v8.21.2: as notificações deixam de bloquear a abertura da Central.
    // A tela principal aparece primeiro; o status de canais chega em segundo plano.
    fetch("/api/notificacoes-jornada/status")
      .then(r => r.json())
      .then(jsonNotif => {
        if (!requisicaoAindaValida(paginaEsperada, tokenEsperado)) return;
        json.notificacoes_jornada = jsonNotif;
        renderAlertasAutomaticos(json);
      })
      .catch(() => {});
  } catch (e) {
    renderCardsDashboardExecutivo({});
    renderAlertasAutomaticos({});
    if (qs("statusCarga")) qs("statusCarga").textContent = "Erro ao carregar Central de Alertas Gerenciais: " + e.message;
  }
}



function prepararFiltrosPrincipaisDaGuia(force = false) {
  // v8.15.24: algumas guias operacionais podem ser abertas depois de telas
  // que escondem ou esvaziam os filtros. Reexibe e repopula os selects
  // independentes do cache de BH 2.0.
  if (qs("filtros")) qs("filtros").classList.remove("hidden");
  if (qs("filtrosDashboard")) qs("filtrosDashboard").classList.add("hidden");
  if (typeof restaurarFiltrosPadraoPrincipal === "function") restaurarFiltrosPadraoPrincipal();

  const ids = ["cc", "nome", "turno"];
  const vazio = ids.some(id => {
    const el = qs(id);
    return el && (!el.options || el.options.length <= 1);
  });

  if (force || vazio) {
    carregarOpcoesFiltros(true);
  }
}

function abrirPagina(pagina, el) {
  // v81738_rota_processamento_manual
  if (pagina === "processamento_manual" || pagina === "processamento_manual_robo") {
    return abrirProcessamentoManualRoboV81737();
  }
  if (pagina === "processamento" || pagina === "processamento_integrado") {
    marcarGuiaAtivaV81738("integrado");
  }
  iniciarTransicaoPagina(pagina);
  ocultarDashboardExecutivoPremium();
  if (pagina !== "dashboard") ocultarPainelBancoHoras2();
  if (pagina !== "extrato_banco_horas") restaurarFiltrosPadraoPrincipal();

  if (pagina === "extrato_banco_horas") {
    paginaAtual = "extrato_banco_horas";
    limparResiduosBancoHoras2ForaDaGuiaBanco();
    dadosConsolidadoAtual = [];
    document.querySelectorAll(".menu button").forEach(b => b.classList.remove("active"));
    if (el) el.classList.add("active");
    carregarExtratoBancoHoras();
    return;
  }


if (pagina === "inter_jornada") {
    paginaAtual = "inter_jornada";
    dadosConsolidadoAtual = [];
    document.querySelectorAll(".menu button").forEach(b => b.classList.remove("active"));
    if (el) el.classList.add("active");
    prepararFiltrosPrincipaisDaGuia(true);
    prepararFiltrosPrincipaisDaGuia(true);
    carregarInterJornada();
    return;
  }

  if (pagina === "violacoes_jornada") {
    paginaAtual = "violacoes_jornada";
    dadosConsolidadoAtual = [];
    document.querySelectorAll(".menu button").forEach(b => b.classList.remove("active"));
    if (el) el.classList.add("active");
    prepararFiltrosPrincipaisDaGuia(true);
    prepararFiltrosPrincipaisDaGuia(true);
    carregarViolacoesJornada();
    return;
  }


ajustarFiltroTipoInconsistencias();
  paginaAtual = pagina;
  renderizarCardsPorGuia();
  marcarPaginaAtualClasse();

  document.querySelectorAll(".menu button").forEach(b => b.classList.remove("active"));
  if (el) el.classList.add("active");

  if (pagina === "dashboard") {
    qs("tituloPagina").textContent = "Banco de Horas";
    qs("subtituloPagina").textContent = "Resumo por colaborador com saldo atual, horas extras e ausências no período.";
    carregarDashboard();

  } else if (pagina === "inconsistencias") {
    qs("tituloPagina").textContent = "Fila de Inconsistências";
    qs("subtituloPagina").textContent = "Exibe apenas registros que exigem tratamento operacional.";
    if (qs("tipo")) qs("tipo").value = "";
    ajustarFiltroTipoInconsistencias();
    prepararFiltrosPrincipaisDaGuia(true);
    carregarOpcoesInconsistencias();
    carregarConsolidado();

  } else if (pagina === "ponto_aberto") {
    paginaAtual = "ponto_aberto";
    renderizarCardsPorGuia();
    qs("tituloPagina").textContent = "Sem Marcação";
    qs("subtituloPagina").textContent = "Colaboradores sem qualquer registro de ponto no período selecionado.";
    if (qs("tipo")) qs("tipo").value = "";
    prepararFiltrosPrincipaisDaGuia(true);
    carregarPontoEmAberto();

  } else if (pagina === "violacoes_jornada") {
    paginaAtual = "violacoes_jornada";
    if (typeof renderizarCardsPorGuia === "function") renderizarCardsPorGuia();
    qs("tituloPagina").textContent = "Violações de Jornada";
    qs("subtituloPagina").textContent = "Alertas de excesso de jornada conforme regra nominal do turno.";
    if (qs("tipo")) qs("tipo").value = "";
    if (qs("filtroInconsistencia")) qs("filtroInconsistencia").value = "";
    carregarViolacoesJornada();

  } else if (pagina === "inter_jornada") {
    paginaAtual = "inter_jornada";
    renderizarCardsPorGuia();
    qs("tituloPagina").textContent = "Violações Inter Jornada";
    qs("subtituloPagina").textContent = "Ocorrências em que o intervalo entre a última marcação e a primeira marcação seguinte é inferior a 11 horas.";
    if (qs("tipo")) qs("tipo").value = "";
    carregarInterJornada();


  } else if (pagina === "ponto_aberto_impar") {
    paginaAtual = "ponto_aberto_impar";
    renderizarCardsPorGuia();
    qs("tituloPagina").textContent = "Ponto Aberto";
    qs("subtituloPagina").textContent = "Registros com quantidade ímpar de marcações, indicando possível batida incompleta.";
    if (qs("tipo")) qs("tipo").value = "";
    prepararFiltrosPrincipaisDaGuia(true);
    carregarPontoAberto();

  } else {
    paginaAtual = "consolidado";
    qs("tituloPagina").textContent = "Espelho de Ponto";
    qs("subtituloPagina").textContent = "Consulta detalhada das marcações, jornada, horas extras, ausências, saldo de banco e validações dos colaboradores.";
    if (qs("tipo")) qs("tipo").value = "";
    if (qs("filtroInconsistencia")) qs("filtroInconsistencia").value = "";
    prepararFiltrosPrincipaisDaGuia(true);
    carregarConsolidado();
  }
}

function filtrarTipo(status, tipo, el) {
  iniciarTransicaoPagina(tipo === "he" ? "hora_extra" : (tipo === "ausencia" ? "ausencias" : "consolidado"));
  ocultarDashboardExecutivoPremium();
  ocultarPainelBancoHoras2();
  if (tipo === "he") {
    paginaAtual = "hora_extra";
    renderizarCardsPorGuia();
    marcarPaginaAtualClasse();
    ajustarControleOrdenacaoBancoHoras();
    renderizarCardsPorGuia();
  } else if (tipo === "ausencia") {
    paginaAtual = "ausencias";
  } else {
    paginaAtual = "consolidado";
  }

  document.querySelectorAll(".menu button").forEach(b => b.classList.remove("active"));
  if (el) el.classList.add("active");

  if (qs("tipo")) qs("tipo").value = tipo || "";
  if (qs("filtroInconsistencia")) qs("filtroInconsistencia").value = "";

  qs("tituloPagina").textContent = el ? el.textContent.trim() : "Visão Filtrada";
  qs("subtituloPagina").textContent =
    tipo === "he"
      ? "Visão exclusiva dos registros com horas extras."
      : tipo === "ausencia"
        ? "Visão exclusiva dos registros com ausências."
        : "Visão filtrada da base consolidada.";

  carregarConsolidado();
}


function atualizarDadosGuiaAtual() {
  // v8.15.25: o botão global Atualizar Dados deve recarregar somente
  // a guia ativa, sem cair sempre no Espelho de Ponto.
  const pagina = paginaAtual || "consolidado";

  if (pagina === "ponto_aberto") return carregarPontoEmAberto();
  if (pagina === "ponto_aberto_impar") return carregarPontoAberto();
  if (pagina === "inter_jornada") return carregarInterJornada();
  if (pagina === "violacoes_jornada") return carregarViolacoesJornada();
  if (pagina === "extrato_banco_horas") return carregarExtratoBancoHoras();
  if (pagina === "dashboard_executivo") return carregarDashboardExecutivo();
  if (pagina === "alertas_automaticos") return carregarAlertasAutomaticos();
  if (pagina === "central_operacao") return carregarCentralOperacaoV8190();
  if (pagina === "diagnostico_instalacao") return carregarDiagnosticoInstalacaoV82128();
  if (pagina === "dashboard" && typeof carregarDashboard === "function") return carregarDashboard();
  if (pagina === "auditoria_premium" && typeof carregarAuditoriaPremium === "function") return carregarAuditoriaPremium();
  if (pagina === "hora_extra" || pagina === "ausencias" || pagina === "inconsistencias" || pagina === "consolidado") return carregarConsolidado();

  return carregarConsolidado();
}

function recarregarPaginaAtualFiltros() {
  // v8.10.45: evita rajadas de chamadas quando select dispara input+change
  // ou quando o usuário troca filtros rapidamente.
  clearTimeout(timerFiltroPrincipal);
  timerFiltroPrincipal = setTimeout(() => {
    if (paginaAtual === "ponto_aberto") carregarPontoEmAberto();
    else if (paginaAtual === "ponto_aberto_impar") carregarPontoAberto();
    else if (paginaAtual === "inter_jornada") carregarInterJornada();
    else if (paginaAtual === "violacoes_jornada") carregarViolacoesJornada();
    else if (paginaAtual === "extrato_banco_horas") carregarExtratoBancoHoras();
    else if (paginaAtual === "dashboard_executivo") carregarDashboardExecutivo();
    else if (paginaAtual === "alertas_automaticos") carregarAlertasAutomaticos();
    else if (paginaAtual === "central_operacao") carregarCentralOperacaoV8190();
    else if (paginaAtual === "diagnostico_instalacao") carregarDiagnosticoInstalacaoV82128();
    else if (paginaAtual === "dashboard" && typeof carregarDashboard === "function") carregarDashboard();
    else carregarConsolidado();
  }, 220);
}














// v8.21.10 - renderizacao progressiva nas guias operacionais que ainda
// montavam tabelas grandes de uma vez, reduzindo travamentos em bases extensas.
function renderInconsistenciasCompacta(rows) {
  const linhas = limitarRenderTabela(aplicarOrdenacao(rows || []), "inconsistencias");
  const paginaRenderizada = paginaAtual;
  const renderLinha = (r) => `
    <tr>
      <td>${r.cc ?? ""}</td>
      <td><strong>${r.nome ?? ""}</strong></td>
      <td>${r.data ?? ""}</td>
      <td>${r.dia ?? ""}</td>
      <td>${r.e1 ?? ""}</td>
      <td>${r.s1 ?? ""}</td>
      <td>${r.e2 ?? ""}</td>
      <td>${r.s2 ?? ""}</td>
      <td>${r.e3 ?? ""}</td>
      <td>${r.s3 ?? ""}</td>
      <td>${r.e4 ?? ""}</td>
      <td>${r.s4 ?? ""}</td>
      <td>${r.jornada ?? ""}</td>
      <td>${r.observacao ?? ""}</td>
      <td>${statusBadge(r.status)}</td>
      <td>${r.inconsistencias ?? ""}</td>
    </tr>
  `;

  atualizarIndicadoresOrdenacao();
  qs("statusCarga").textContent = linhas.length ? `Renderizando ${linhas.length} inconsistencia(s)...` : "Nenhuma inconsistencia encontrada.";
  renderTabelaEmLotes("tbody", linhas, renderLinha, {
    aoConcluir: () => {
      if (paginaAtual === paginaRenderizada) qs("statusCarga").textContent = `${linhas.length} inconsistencia(s) carregada(s).`;
    },
  });
}

function renderSemMarcacaoCompacta(rows) {
  const linhas = limitarRenderTabela(rows || [], "registros sem marcacao");
  const paginaRenderizada = paginaAtual;
  const renderLinha = (r) => `
    <tr>
      <td>${r.cc ?? ""}</td>
      <td><strong>${r.nome ?? ""}</strong></td>
      <td>${r.funcao ?? ""}</td>
      <td>${r.turno ?? ""}</td>
      <td>${r.data ?? ""}</td>
      <td>${r.dia ?? ""}</td>
      <td>${r.e1 ?? ""}</td>
      <td>${r.s1 ?? ""}</td>
      <td>${r.e2 ?? ""}</td>
      <td>${r.s2 ?? ""}</td>
      <td>${r.e3 ?? ""}</td>
      <td>${r.s3 ?? ""}</td>
      <td>${r.e4 ?? ""}</td>
      <td>${r.s4 ?? ""}</td>
      <td>${r.abono ?? ""}</td>
      <td>${r.absent ? badge(r.absent, "danger") : ""}</td>
      <td>${r.observacao ?? ""}</td>
    </tr>
  `;

  qs("statusCarga").textContent = linhas.length ? `Renderizando ${linhas.length} registro(s) sem marcacao...` : "Nenhum registro sem marcacao encontrado.";
  renderTabelaEmLotes("tbody", linhas, renderLinha, {
    aoConcluir: () => {
      if (paginaAtual === paginaRenderizada) qs("statusCarga").textContent = `${linhas.length} registro(s) sem marcacao carregado(s).`;
    },
  });
}

function renderPontoAbertoCompacto(rows) {
  const linhas = limitarRenderTabela(somenteImparesPontoAberto(rows || []), "registros com ponto aberto");
  const paginaRenderizada = paginaAtual;
  const renderLinha = (r) => `
    <tr>
      <td>${r.cc ?? ""}</td>
      <td><strong>${r.nome ?? ""}</strong></td>
      <td>${r.funcao ?? ""}</td>
      <td>${r.turno ?? ""}</td>
      <td>${r.data ?? ""}</td>
      <td>${r.dia ?? ""}</td>
      <td>${r.e1 ?? ""}</td>
      <td>${r.s1 ?? ""}</td>
      <td>${r.e2 ?? ""}</td>
      <td>${r.s2 ?? ""}</td>
      <td>${r.e3 ?? ""}</td>
      <td>${r.s3 ?? ""}</td>
      <td>${r.e4 ?? ""}</td>
      <td>${r.s4 ?? ""}</td>
      <td>${r.absent ? badge(r.absent, "danger") : ""}</td>
      <td>${r.observacao ?? ""}</td>
      <td>${statusBadge(r.status)}</td>
      <td>${r.inconsistencias ?? ""}</td>
    </tr>
  `;

  const nomes = new Set(linhas.map(r => String(r.nome || "").trim()).filter(Boolean));
  if (qs("kpiRegistros")) qs("kpiRegistros").textContent = linhas.length;
  if (qs("kpiColaboradores")) qs("kpiColaboradores").textContent = nomes.size;
  if (qs("kpiHE")) qs("kpiHE").textContent = "00:00";
  if (qs("kpiAbsent")) qs("kpiAbsent").textContent = "00:00";
  if (qs("kpiSaldoNegativo")) qs("kpiSaldoNegativo").textContent = "0";
  if (qs("kpiRevisar")) qs("kpiRevisar").textContent = linhas.length;

  qs("statusCarga").textContent = linhas.length ? `Renderizando ${linhas.length} registro(s) com marcacoes impares...` : "Nenhum ponto aberto encontrado.";
  renderTabelaEmLotes("tbody", linhas, renderLinha, {
    aoConcluir: () => {
      if (paginaAtual === paginaRenderizada) qs("statusCarga").textContent = `${linhas.length} registro(s) com marcacoes impares carregado(s).`;
    },
  });
}

function renderHoraExtraCompacta(rows) {
  const linhas = limitarRenderTabela(aplicarOrdenacao(rows || []), "registros de hora extra");
  const paginaRenderizada = paginaAtual;
  const renderLinha = (r) => `
    <tr>
      <td>${r.cc ?? ""}</td>
      <td><strong>${r.nome ?? ""}</strong></td>
      <td>${r.funcao ?? ""}</td>
      <td>${r.turno ?? ""}</td>
      <td>${r.data ?? ""}</td>
      <td>${r.dia ?? ""}</td>
      <td>${r.e1 ?? ""}</td>
      <td>${r.s1 ?? ""}</td>
      <td>${r.e2 ?? ""}</td>
      <td>${r.s2 ?? ""}</td>
      <td>${r.e3 ?? ""}</td>
      <td>${r.s3 ?? ""}</td>
      <td>${r.e4 ?? ""}</td>
      <td>${r.s4 ?? ""}</td>
      <td>${r.he ? badge(r.he, "warning") : ""}</td>
      <td>${r.jornada ?? ""}</td>
      <td>${r.observacao ?? ""}</td>
      <td>${statusBadge(r.status)}</td>
      <td>${r.inconsistencias ?? ""}</td>
    </tr>
  `;

  atualizarIndicadoresOrdenacao();
  qs("statusCarga").textContent = linhas.length ? `Renderizando ${linhas.length} registro(s) de hora extra...` : "Nenhum registro de hora extra encontrado.";
  renderTabelaEmLotes("tbody", linhas, renderLinha, {
    aoConcluir: () => {
      if (paginaAtual === paginaRenderizada) qs("statusCarga").textContent = `${linhas.length} registro(s) de hora extra carregado(s).`;
    },
  });
}

function renderInterJornada(rows) {
  const linhas = limitarRenderTabela(rows || [], "interjornadas");
  const paginaRenderizada = paginaAtual;
  const renderLinha = (r) => `
    <tr>
      <td>${r.cc ?? ""}</td>
      <td><strong>${r.nome ?? ""}</strong></td>
      <td>${r.funcao ?? ""}</td>
      <td>${r.turno ?? ""}</td>
      <td>${r.data_anterior ?? ""}</td>
      <td>${r.ultima_marcacao_anterior ?? ""}</td>
      <td>${r.data_retorno ?? ""}</td>
      <td>${r.primeira_marcacao_retorno ?? ""}</td>
      <td>${r.intervalo_interjornada ?? ""}</td>
      <td>${r.deficit_interjornada ?? ""}</td>
      <td>${r.observacao ?? ""}</td>
    </tr>
  `;

  if (qs("statusCarga")) qs("statusCarga").textContent = linhas.length ? `Renderizando ${linhas.length} ocorrencia(s) inter jornada...` : "Nenhuma violacao inter jornada encontrada.";
  renderTabelaEmLotes("tbody", linhas, renderLinha, {
    aoConcluir: () => {
      if (paginaAtual === paginaRenderizada && qs("statusCarga")) qs("statusCarga").textContent = `${linhas.length} ocorrencia(s) de violacoes inter jornada carregada(s).`;
    },
  });
}

function renderViolacoesJornada(rows) {
  const linhas = limitarRenderTabela(rows || [], "violacoes");
  const paginaRenderizada = paginaAtual;
  const renderLinha = (r) => `
    <tr>
      <td>${r.cc ?? ""}</td>
      <td><strong>${r.nome ?? ""}</strong></td>
      <td>${r.funcao ?? ""}</td>
      <td>${r.tipo_jornada ?? ""}</td>
      <td>${r.periodo ?? ""}</td>
      <td>${r.data_referencia ?? ""}</td>
      <td>${r.violacao ?? ""}</td>
      <td>${r.detalhe ?? ""}</td>
      <td>${r.dias_trabalhados ?? ""}</td>
      <td>${r.limite_esperado ?? ""}</td>
    </tr>
  `;

  if (qs("statusCarga")) qs("statusCarga").textContent = linhas.length ? `Renderizando ${linhas.length} violacao(oes) de jornada...` : "Nenhuma violacao de jornada encontrada.";
  renderTabelaEmLotes("tbody", linhas, renderLinha, {
    aoConcluir: () => {
      if (paginaAtual === paginaRenderizada && qs("statusCarga")) qs("statusCarga").textContent = `${linhas.length} violacao(oes) de jornada carregada(s).`;
    },
  });
}

function renderExtratoBancoHoras(rows) {
  const linhas = limitarRenderTabela(rows || [], "linhas de extrato");
  const paginaRenderizada = paginaAtual;
  const renderLinha = (r) => `
    <tr>
      <td>${r.data ?? ""}</td>
      <td>${r.data_emissao_pdf ?? ""}</td>
      <td><strong>${r.nome ?? ""}</strong></td>
      <td>${r.cc ?? ""}</td>
      <td>${r.turno ?? ""}</td>
      <td>${r.dia ?? ""}</td>
      <td><strong>${r.evento ?? ""}</strong></td>
      <td>${r.he ?? ""}</td>
      <td>${r.absent ?? ""}</td>
      <td><strong>${r.saldo_calculado ?? ""}</strong></td>
      <td>${r.saldo_informado ?? r.saldo_pdf ?? ""}</td>
      <td>${r.observacao ?? ""}</td>
      <td>${r.inconsistencias ?? ""}</td>
    </tr>
  `;

  if (qs("statusCarga")) qs("statusCarga").textContent = linhas.length ? `Renderizando ${linhas.length} linha(s) de extrato...` : "Nenhuma linha de extrato encontrada.";
  renderTabelaEmLotes("tbody", linhas, renderLinha, {
    aoConcluir: () => {
      if (paginaAtual === paginaRenderizada && qs("statusCarga")) qs("statusCarga").textContent = `${linhas.length} linha(s) de extrato carregada(s).`;
    },
  });
}

function limparFiltros() {
  ["turno", "nome", "data", "tipoInconsistencia", "tipo", "filtroInconsistencia"].forEach(id => {
    if (qs(id)) qs(id).value = "";
  });
  limparDatasFiltro();
  ordenacaoSaldo = null;
  ordenacaoHE = null;
  ordenacaoInconsistencias = null;
  recarregarPaginaAtualFiltros();
}

function limparFiltrosDashboard() {
  if (qs("ccDashboard")) qs("ccDashboard").value = "";
  if (qs("turnoDashboard")) qs("turnoDashboard").value = "";
  if (qs("nomeDashboard")) qs("nomeDashboard").value = "";
  if (qs("dataDashboard")) qs("dataDashboard").value = "";
  if (qs("ordemSaldoBanco")) qs("ordemSaldoBanco").value = "";
  limparDatasDashboardFiltro();
  carregarDashboard();
}

function zerarKpis() {
  if (qs("kpiRegistros")) qs("kpiRegistros").textContent = "0";
  if (qs("kpiColaboradores")) qs("kpiColaboradores").textContent = "0";
  if (qs("kpiHE")) qs("kpiHE").textContent = "00:00";
  if (qs("kpiAbsent")) qs("kpiAbsent").textContent = "00:00";
  if (qs("kpiSaldoNegativo")) qs("kpiSaldoNegativo").textContent = "0";
  if (qs("kpiRevisar")) qs("kpiRevisar").textContent = "0";
  if (qs("kpiBancoPositivo")) qs("kpiBancoPositivo").textContent = "00,00";
  if (qs("kpiBancoNegativoTotal")) qs("kpiBancoNegativoTotal").textContent = "00,00";
  if (qs("kpiColabBancoPositivo")) qs("kpiColabBancoPositivo").textContent = "0";
}



async function carregarStatusBaseInicial() {
  try {
    const resp = await fetch("/api/status-base");
    const json = await resp.json();
    if (!json.base_ativa) {
      if (qs("statusCarga")) qs("statusCarga").textContent = json.mensagem;
      if (qs("tbody")) qs("tbody").innerHTML = "";
    }
  } catch (e) {}
}


document.addEventListener("DOMContentLoaded", () => {
  ["turno", "nome", "data", "tipoInconsistencia", "tipo", "filtroInconsistencia"].forEach(id => {
    const el = qs(id);
    if (el && !el.dataset.listenerPrincipalV81045) {
      el.dataset.listenerPrincipalV81045 = "1";
      el.addEventListener("change", recarregarPaginaAtualFiltros);
    }
  });

  ["ccDashboard", "nomeDashboard", "turnoDashboard", "dataDashboard", "ordemSaldoBanco"].forEach(id => {
    const el = qs(id);
    if (el && !el.dataset.listenerDashboardV81045) {
      el.dataset.listenerDashboardV81045 = "1";
      el.addEventListener("change", carregarDashboard);
    }
  });

  carregarUsuarioAtual();
  carregarOpcoesFiltros(true);
  vincularEventoComboCC();
  renderDatasSelecionadas();
  renderDatasDashboardSelecionadas();
  carregarConsolidado();
});


document.addEventListener('DOMContentLoaded', carregarStatusBaseInicial);


// v8.10.45: listeners duplicados removidos.
// O recarregamento de filtros agora passa por recarregarPaginaAtualFiltros(),
// com debounce e token de navegação, evitando múltiplas chamadas simultâneas.



async function fetchJsonSeguro(url, options = {}) {
  const resp = await fetch(url, options);
  const texto = await resp.text();

  let json = null;
  try {
    json = texto ? JSON.parse(texto) : {};
  } catch (e) {
    const limpo = (texto || "").replace(/<[^>]*>/g, " ").replace(/\s+/g, " ").trim().slice(0, 300);
    if (resp.status === 401 || limpo.toLowerCase().includes("login")) {
      window.location.href = "/login";
      throw new Error("Sessão expirada. Faça login novamente.");
    }
    throw new Error(`Servidor devolveu HTML em vez de JSON (${resp.status}). ${limpo || texto.slice(0, 180)}`);
  }

  if (!resp.ok || json.ok === false || json.erro) {
    throw new Error(json.erro || json.mensagem || `HTTP ${resp.status}`);
  }
  return json;
}


// ============================================================
// PROCESSAMENTO INTEGRADO - v8.11.0
// ============================================================
let timerProcessamentoIntegrado = null;

function renderProcessamentoStatus(job) {
  const alvo = qs("processamentoStatusBox");
  if (!alvo) return;
  job = job || {};
  const status = job.status || "ocioso";
  const ok = job.ok !== false;
  const classe = status === "rodando" ? "warning" : (ok ? "success" : "danger");
  const resultado = job.resultado || {};
  const launcher = resultado.launcher || {};
  const imp = resultado.importacao_excel || {};
  const hist = resultado.historico_pdfs || {};
  const bh = resultado.banco_horas || {};
  alvo.innerHTML = `
    <div class="process-status-card ${classe}">
      <div>
        <span class="exec-eyebrow">Status</span>
        <h4>${escapeHtml(job.fase || status)}</h4>
        <p>${escapeHtml(job.mensagem || "")}</p>
      </div>
      <div class="process-status-meta">
        <strong>${escapeHtml(status)}</strong>
        <span>Início: ${escapeHtml(job.iniciado_em || "-")}</span>
        <span>Fim: ${escapeHtml(job.finalizado_em || "-")}</span>
        <span>Tempo: ${escapeHtml(String(job.segundos || 0))}s</span>
      </div>
    </div>
    <div class="process-mini-grid">
      <div><strong>Launcher</strong><span>${launcher.ok === undefined ? "Aguardando" : (launcher.ok ? "OK" : "Erro")}</span><small>${escapeHtml(launcher.caminho || "")}</small></div>
      <div><strong>Excel</strong><span>${imp.ok === undefined ? "Aguardando" : (imp.ok ? "Importado" : "Erro")}</span><small>${escapeHtml(imp.mensagem || "")}</small></div>
      <div><strong>PDFs</strong><span>${hist.quadros_lidos !== undefined ? `${hist.quadros_lidos} quadro(s)` : "Opcional"}</span><small>${escapeHtml(hist.mensagem || "")}</small></div>
      <div><strong>Banco de Horas</strong><span>${bh.ok === undefined ? "Opcional" : (bh.ok ? "Atualizado" : "Erro")}</span><small>${escapeHtml(bh.mensagem || "")}</small></div>
    </div>
    <details class="process-log-box">
      <summary>Ver log</summary>
      <pre>${escapeHtml(job.log_tail || "Sem log registrado ainda.")}</pre>
    </details>
  `;
}

async function consultarStatusProcessamentoIntegrado() {
  try {
    // v8.11.2: status do job pode ser erro, mas a consulta do status não deve quebrar a tela.
    // Antes usávamos fetchJsonSeguro(), que tratava job.ok=false como falha de requisição
    // e escondia justamente o log técnico necessário para diagnóstico.
    const resp = await fetch("/api/processamento-integrado/status");
    const job = await resp.json();
    if (!resp.ok || job.erro) throw new Error(job.erro || job.mensagem || `HTTP ${resp.status}`);
    renderProcessamentoStatus(job);
    if (job.status !== "rodando" && timerProcessamentoIntegrado) {
      clearInterval(timerProcessamentoIntegrado);
      timerProcessamentoIntegrado = null;
      if (job.status === "concluido") {
        if (qs("statusCarga")) qs("statusCarga").textContent = "Processamento integrado concluído.";
      } else if (job.status === "erro") {
        if (qs("statusCarga")) qs("statusCarga").textContent = "Processamento integrado terminou com erro. Veja o log técnico.";
      }
    }
    return job;
  } catch (e) {
    const alvo = qs("processamentoStatusBox");
    if (alvo) alvo.innerHTML = `<div class="exec-empty erro">Falha ao consultar status: ${escapeHtml(e.message)}</div>`;
  }
}


async function iniciarProcessamentoIntegrado() {
  if (!usuarioAtualEhAdmin()) {
    alert("Acesso negado. Esta ação exige perfil Administrador.");
    return;
  }
  const payload = {
    // v8.11.9: o Processamento Integrado volta a ser uma etapa segura e simples:
    // executar o Launcher externo e importar o Excel gerado. PDFs/BH ficam concentrados
    // exclusivamente em Configuração do Robô para não contaminar o núcleo Banco/Extrato.
    importar_excel: true,
    atualizar_historico_pdfs: false,
    corrigir_banco_horas: false,
    timeout_segundos: 3600,
  };
  try {
    const json = await fetchJsonSeguro("/api/processamento-integrado/iniciar", {
      method: "POST",
      headers: { "Content-Type": "application/json", "Accept": "application/json" },
      credentials: "same-origin",
      body: JSON.stringify(payload),
    });
    renderProcessamentoStatus(json.job || json);
    if (qs("statusCarga")) qs("statusCarga").textContent = "Processamento integrado iniciado em segundo plano.";
    if (timerProcessamentoIntegrado) clearInterval(timerProcessamentoIntegrado);
    timerProcessamentoIntegrado = setInterval(consultarStatusProcessamentoIntegrado, 2500);
    setTimeout(consultarStatusProcessamentoIntegrado, 800);
    setTimeout(carregarPainelCicloOperacional, 1200);
    setTimeout(carregarHistoricoProcessamento, 1600);
  } catch (e) {
    alert("Falha ao iniciar processamento integrado: " + e.message);
  }
}

function lerOpcoesCleanupOperacional() {
  const marcado = (id, padrao = true) => {
    const el = qs(id);
    return el ? !!el.checked : padrao;
  };
  return {
    limpar_base_app: marcado("cleanupBaseApp", true),
    limpar_cache_bh: marcado("cleanupCacheBh", true),
    limpar_pdfs_entrada: marcado("cleanupEntrada", true),
    limpar_pdfs_processados: marcado("cleanupProcessados", false),
    limpar_pdfs_erro: marcado("cleanupErro", false),
    limpar_saida_robo: marcado("cleanupSaida", true),
    limpar_controle_temporario: marcado("cleanupControle", true),
    limpar_logs_historico: marcado("cleanupLogsHistorico", false),
    limpar_pastas_pdf: false,
  };
}

function renderResultadoCleanupOperacional(json) {
  const alvo = qs("processamentoStatusBox");
  if (!alvo) return;
  const removidos = json.removidos || [];
  const erros = json.erros || [];
  alvo.innerHTML = `
    <div class="process-status-card ${erros.length ? "warning" : "success"}">
      <div>
        <span class="exec-eyebrow">Cleanup</span>
        <h4>${escapeHtml(json.mensagem || "Cleanup concluido")}</h4>
        <p>${escapeHtml(String(json.total_removidos || 0))} arquivo(s) removido(s). Configuracoes, usuarios e auditoria preservados.</p>
      </div>
      <div class="process-status-meta">
        <strong>${erros.length ? "PARCIAL" : "OK"}</strong>
        <span>${escapeHtml(String(Math.round((json.bytes || 0) / 1024)))} KB liberados</span>
      </div>
    </div>
    <details class="process-log-box" open>
      <summary>Itens removidos</summary>
      <pre>${escapeHtml(removidos.slice(0, 120).map(x => `${x.categoria}: ${x.caminho}`).join("\n") || "Nenhum arquivo operacional encontrado para remover.")}</pre>
    </details>
    ${erros.length ? `<details class="process-log-box" open><summary>Itens nao removidos</summary><pre>${escapeHtml(erros.map(x => `${x.categoria}: ${x.caminho} - ${x.erro}`).join("\n"))}</pre></details>` : ""}
  `;
}

async function executarCleanupOperacional() {
  if (!usuarioAtualEhAdmin()) {
    alert("Acesso negado. Esta acao exige perfil Administrador.");
    return;
  }
  const opcoes = lerOpcoesCleanupOperacional();
  const itens = [
    opcoes.limpar_base_app ? "tabelas/base carregada do Tempo Fechado" : null,
    opcoes.limpar_cache_bh ? "caches de Banco de Horas e PDFs" : null,
    opcoes.limpar_pdfs_entrada ? "PDFs da entrada" : null,
    opcoes.limpar_pdfs_processados ? "PDFs processados" : null,
    opcoes.limpar_pdfs_erro ? "PDFs em erro" : null,
    opcoes.limpar_saida_robo ? "Excel/CSV da saida do robo" : null,
    opcoes.limpar_controle_temporario ? "temporarios de controle" : null,
    opcoes.limpar_logs_historico ? "Historico de Processamentos e logs" : null,
  ].filter(Boolean);
  if (!itens.length) {
    alert("Selecione ao menos um item para limpar.");
    return;
  }
  const msg = [
    "Executar Cleanup Operacional?",
    "",
    "Serao limpos:",
    ...itens.map(x => "- " + x),
    "",
    opcoes.limpar_logs_historico
      ? "Atencao: Historico de Processamentos e logs serao apagados para instalacao limpa."
      : "Configuracoes, usuarios, auditoria e logs principais serao preservados."
  ].join("\n");
  if (!confirm(msg)) return;

  try {
    if (qs("statusCarga")) qs("statusCarga").textContent = "Executando Cleanup Operacional...";
    const json = await fetchJsonSeguro("/api/cleanup-operacional", {
      method: "POST",
      headers: { "Content-Type": "application/json", "Accept": "application/json" },
      credentials: "same-origin",
      body: JSON.stringify(opcoes),
    });

    dadosConsolidadoAtual = [];
    limparCacheOpcoesFiltros();
    if (qs("tbody")) qs("tbody").innerHTML = "";
    if (qs("tbodyDashboard")) qs("tbodyDashboard").innerHTML = "";
    if (qs("dashboardExecutivoPremium")) qs("dashboardExecutivoPremium").classList.remove("hidden");
    zerarKpis();
    if (qs("statusCarga")) qs("statusCarga").textContent = json.mensagem || "Cleanup concluido.";
    renderResultadoCleanupOperacional(json);
    await Promise.allSettled([carregarPainelCicloOperacional(), carregarHistoricoProcessamento(), executarPrevalidacaoProcessamento(false)]);
    alert(json.mensagem || "Cleanup concluido.");
  } catch (e) {
    alert("Falha ao executar Cleanup Operacional: " + e.message);
    if (qs("statusCarga")) qs("statusCarga").textContent = "Falha no Cleanup Operacional.";
  }
}

function renderPainelCicloOperacional(json) {
  const alvo = qs("cicloOperacionalBox");
  if (!alvo) return;
  const base = json.base_app || {};
  const robo = json.robo || {};
  const job = json.job || {};
  const cacheOk = Object.values(json.caches || {}).filter(x => x && x.existe).length;
  const etapa = job.status === "rodando" ? (job.fase || "Rodando") : (job.status || "ocioso");
  alvo.innerHTML = `
    <div class="process-mini-grid">
      <div><strong>Base do app</strong><span>${base.excel_existe ? "Carregada" : "Vazia"}</span><small>${escapeHtml((base.excel_app || {}).idade || base.arquivo || "-")}</small></div>
      <div><strong>Entrada</strong><span>${escapeHtml(String(((robo.entrada || {}).pdfs) || 0))} PDF(s)</span><small>${escapeHtml((robo.entrada || {}).caminho || "-")}</small></div>
      <div><strong>Saida</strong><span>${escapeHtml(String(((robo.saida || {}).planilhas) || 0))} arquivo(s)</span><small>${escapeHtml((robo.excel || {}).idade || (robo.saida || {}).caminho || "-")}</small></div>
      <div><strong>Erros</strong><span>${escapeHtml(String(((robo.erro || {}).pdfs) || 0))} PDF(s)</span><small>${escapeHtml((robo.erro || {}).caminho || "-")}</small></div>
      <div><strong>Caches BH</strong><span>${cacheOk} ativo(s)</span><small>PDFs, resumo e calculo</small></div>
      <div><strong>Processo</strong><span>${escapeHtml(etapa)}</span><small>${escapeHtml(job.mensagem || json.gerado_em || "")}</small></div>
    </div>
  `;
}

async function carregarPainelCicloOperacional() {
  const alvo = qs("cicloOperacionalBox");
  if (alvo) alvo.innerHTML = `<div class="exec-empty">Atualizando ciclo operacional...</div>`;
  try {
    const json = await fetchJsonSeguro("/api/ciclo-operacional");
    renderPainelCicloOperacional(json);
  } catch (e) {
    if (alvo) alvo.innerHTML = `<div class="exec-empty erro">Falha ao carregar ciclo: ${escapeHtml(e.message)}</div>`;
  }
}

function renderPrevalidacaoProcessamento(json) {
  const alvo = qs("prevalidacaoBox");
  if (!alvo) return;
  const checks = json.checks || [];
  const nivel = json.nivel === "erro" ? "danger" : (json.nivel === "warning" ? "warning" : "success");
  alvo.innerHTML = `
    <div class="process-status-card ${nivel}">
      <div>
        <span class="exec-eyebrow">Pre-validacao</span>
        <h4>${json.pronto ? "Ambiente pronto" : "Revisar antes de coletar"}</h4>
        <p>${escapeHtml(json.mensagem || "")}</p>
      </div>
      <div class="process-status-meta">
        <strong>${escapeHtml((json.nivel || "ok").toUpperCase())}</strong>
        <span>${escapeHtml(String((json.resumo || {}).avisos || 0))} aviso(s)</span>
        <span>${escapeHtml(String((json.resumo || {}).erros || 0))} erro(s)</span>
      </div>
    </div>
    <div class="process-mini-grid">
      ${checks.map(c => `<div><strong>${escapeHtml(c.nome || "-")}</strong><span>${c.ok ? "OK" : (c.nivel === "warning" ? "Aviso" : "Erro")}</span><small>${escapeHtml(c.detalhe || "")}</small></div>`).join("")}
    </div>
  `;
}

async function executarPrevalidacaoProcessamento(mostrarStatus = true) {
  const alvo = qs("prevalidacaoBox");
  if (mostrarStatus && alvo) alvo.innerHTML = `<div class="exec-empty">Executando pre-validacao...</div>`;
  try {
    const json = await fetchJsonSeguro("/api/processamento/prevalidacao");
    renderPrevalidacaoProcessamento(json);
    return json;
  } catch (e) {
    if (alvo) alvo.innerHTML = `<div class="exec-empty erro">Falha na pre-validacao: ${escapeHtml(e.message)}</div>`;
    return null;
  }
}

function renderHistoricoProcessamento(json) {
  const alvo = qs("historicoProcessamentoBox");
  if (!alvo) return;
  const eventos = json.historico || [];
  alvo.innerHTML = eventos.length ? `
    <div class="process-mini-grid">
      ${eventos.slice(0, 12).map(e => `<div><strong>${escapeHtml(e.operacao || e.nome || "Operacao")}</strong><span>${escapeHtml(e.status || "-")}</span><small>${escapeHtml((e.quando || e.data || "") + " " + (e.mensagem || ""))}</small></div>`).join("")}
    </div>
  ` : `<div class="exec-empty">Nenhum historico operacional registrado ainda.</div>`;
}

async function carregarHistoricoProcessamento() {
  const alvo = qs("historicoProcessamentoBox");
  if (alvo) alvo.innerHTML = `<div class="exec-empty">Carregando historico...</div>`;
  try {
    const json = await fetchJsonSeguro("/api/processamento/historico?limite=20");
    renderHistoricoProcessamento(json);
  } catch (e) {
    if (alvo) alvo.innerHTML = `<div class="exec-empty erro">Falha ao carregar historico: ${escapeHtml(e.message)}</div>`;
  }
}

async function executarNovaRodadaOperacional() {
  if (!usuarioAtualEhAdmin()) {
    alert("Acesso negado. Esta acao exige perfil Administrador.");
    return;
  }
  const comCleanup = qs("novaRodadaComCleanup") ? qs("novaRodadaComCleanup").checked : true;
  if (!confirm(comCleanup ? "Iniciar Nova Rodada com Cleanup previo e coleta?" : "Iniciar Nova Rodada sem Cleanup previo?")) return;
  try {
    if (qs("statusCarga")) qs("statusCarga").textContent = "Iniciando Nova Rodada...";
    const json = await fetchJsonSeguro("/api/processamento/nova-rodada", {
      method: "POST",
      headers: { "Content-Type": "application/json", "Accept": "application/json" },
      credentials: "same-origin",
      body: JSON.stringify({
        executar_cleanup: comCleanup,
        cleanup: lerOpcoesCleanupOperacional(),
        timeout_segundos: 3600,
      }),
    });
    renderProcessamentoStatus(json.job || json);
    if (timerProcessamentoIntegrado) clearInterval(timerProcessamentoIntegrado);
    timerProcessamentoIntegrado = setInterval(async () => {
      await consultarStatusProcessamentoIntegrado();
      await carregarPainelCicloOperacional();
    }, 2500);
    await Promise.allSettled([carregarPainelCicloOperacional(), carregarHistoricoProcessamento(), executarPrevalidacaoProcessamento(false)]);
  } catch (e) {
    alert("Falha ao iniciar Nova Rodada: " + e.message);
    if (qs("statusCarga")) qs("statusCarga").textContent = "Falha ao iniciar Nova Rodada.";
  }
}

function montarHtmlProcessamentoIntegradoV82117() {
  return `
    <div class="process-toolbar">
      <div class="process-toolbar-title">
        <h3>Processamento Integrado</h3>
      </div>
      <div class="process-toolbar-actions">
        <button class="btn" onclick="iniciarProcessamentoIntegrado()">Executar coleta</button>
        <button class="btn" onclick="executarNovaRodadaOperacional()">Nova Rodada</button>
        <button class="btn secondary" onclick="executarPrevalidacaoProcessamento()">Pre-validar</button>
        <button class="btn secondary" onclick="executarCleanupOperacional()">Cleanup</button>
        <button class="btn secondary" onclick="Promise.allSettled([consultarStatusProcessamentoIntegrado(), carregarPainelCicloOperacional(), carregarHistoricoProcessamento(), executarPrevalidacaoProcessamento(false)])">Atualizar</button>
        <button class="btn secondary" onclick="abrirConfiguracaoRobo()">Configurar</button>
      </div>
    </div>
    <section class="exec-panel process-panel-compact">
      <div class="exec-panel-title"><div><h4>Painel de Ciclo</h4><p>Base carregada, pastas do robo, caches e processo atual.</p></div></div>
      <div id="cicloOperacionalBox" class="processamento-status-box"><div class="exec-empty">Carregando ciclo operacional...</div></div>
    </section>
    <section class="exec-panel process-panel-compact">
      <div class="exec-panel-title"><div><h4>Nova Rodada e Cleanup</h4><p>Escolha exatamente o que limpar antes de uma nova avaliacao.</p></div></div>
      <div class="process-mini-grid cleanup-options-grid">
        <label><input id="novaRodadaComCleanup" type="checkbox" checked> Cleanup antes da Nova Rodada</label>
        <label><input id="cleanupBaseApp" type="checkbox" checked> Tabelas/base do app</label>
        <label><input id="cleanupCacheBh" type="checkbox" checked> Caches Banco de Horas</label>
        <label><input id="cleanupEntrada" type="checkbox" checked> PDFs entrada</label>
        <label><input id="cleanupProcessados" type="checkbox"> PDFs processados</label>
        <label><input id="cleanupErro" type="checkbox"> PDFs erro</label>
        <label><input id="cleanupSaida" type="checkbox" checked> Excel/CSV saida</label>
        <label><input id="cleanupControle" type="checkbox" checked> Temporarios controle</label>
        <label><input id="cleanupLogsHistorico" type="checkbox"> Historico de Processamentos e logs</label>
      </div>
    </section>
    <section class="exec-panel process-panel-compact">
      <div class="exec-panel-title"><div><h4>Pre-validacao</h4><p>Checklist rapido antes da coleta.</p></div></div>
      <div id="prevalidacaoBox" class="processamento-status-box"><div class="exec-empty">Aguardando pre-validacao...</div></div>
    </section>
    <section class="exec-panel process-panel-compact">
      <div class="exec-panel-title"><div><h4>Andamento</h4><p>Ultima execucao e log.</p></div></div>
      <div id="processamentoStatusBox" class="processamento-status-box"><div class="exec-empty">Consultando status...</div></div>
    </section>
    <section class="exec-panel process-panel-compact">
      <div class="exec-panel-title"><div><h4>Historico de Processamentos</h4><p>Ultimos eventos operacionais registrados.</p></div></div>
      <div id="historicoProcessamentoBox" class="processamento-status-box"><div class="exec-empty">Carregando historico...</div></div>
    </section>
  `;
}

async function abrirProcessamentoIntegrado(btn) {
  marcarGuiaAtivaV81738("integrado");
  limparEstadoOperacionalV81738();
  iniciarTransicaoPagina("processamento_integrado");
  if (!usuarioAtualEhAdmin()) {
    alert("Acesso negado. Esta area exige perfil Administrador.");
    return;
  }
  paginaAtual = "processamento_integrado";
  document.querySelectorAll(".menu button").forEach(b => b.classList.remove("active"));
  if (btn) btn.classList.add("active");
  marcarPaginaAtualClasse();
  setVisibilidadeDashboardExecutivoPremium(true);
  if (qs("filtros")) qs("filtros").classList.add("hidden");
  if (qs("filtrosDashboard")) qs("filtrosDashboard").classList.add("hidden");
  if (qs("cardsContainer")) qs("cardsContainer").classList.add("hidden");
  if (qs("controleOrdenacaoBancoHoras")) qs("controleOrdenacaoBancoHoras").classList.add("hidden");
  if (qs("tituloPagina")) qs("tituloPagina").textContent = "Processamento";
  if (qs("subtituloPagina")) qs("subtituloPagina").textContent = "Coleta e importacao assistida do Tempo Fechado.";
  if (qs("tituloTabela")) qs("tituloTabela").textContent = "";
  if (qs("statusCarga")) qs("statusCarga").textContent = "Aguardando comando de processamento.";

  const painel = qs("dashboardExecutivoPremium");
  if (painel) {
    painel.classList.remove("hidden");
    painel.innerHTML = montarHtmlProcessamentoIntegradoV82117();
  }
  await Promise.allSettled([
    consultarStatusProcessamentoIntegrado(),
    carregarPainelCicloOperacional(),
    executarPrevalidacaoProcessamento(false),
    carregarHistoricoProcessamento(),
  ]);
}

async function abrirAdministracao(btn) {
  iniciarTransicaoPagina("administracao");
  if (!usuarioAtualEhAdmin()) {
    alert("Acesso negado. Esta área exige perfil Administrador.");
    return;
  }
  paginaAtual = "administracao";
  document.querySelectorAll(".menu button").forEach(b => b.classList.remove("active"));
  if (btn) btn.classList.add("active");

  qs("tituloPagina").textContent = "Administração Multiusuário";
  qs("subtituloPagina").textContent = "Usuários, permissões e auditoria.";
  marcarPaginaAtualClasse();
  setVisibilidadeDashboardExecutivoPremium(true);
  if (qs("filtros")) qs("filtros").classList.add("hidden");
  if (qs("filtrosDashboard")) qs("filtrosDashboard").classList.add("hidden");
  if (qs("cardsContainer")) qs("cardsContainer").classList.add("hidden");
  if (qs("controleOrdenacaoBancoHoras")) qs("controleOrdenacaoBancoHoras").classList.add("hidden");
  const painel = qs("dashboardExecutivoPremium");
  if (painel) {
    painel.classList.remove("hidden");
    painel.innerHTML = `<div class="admin-panel"><h3>Carregando auditoria...</h3></div>`;
  }
  await carregarAuditoriaMultiusuario();
}

async function carregarAuditoriaMultiusuario() {
  const painel = qs("dashboardExecutivoPremium");
  const paginaEsperada = "administracao";
  const tokenEsperado = navegacaoToken;
  try {
    const [jsonUsuarios, jsonAuditoria] = await Promise.all([
      fetchJsonSeguro("/api/usuarios"),
      fetchJsonSeguro("/api/auditoria"),
    ]);
    if (!requisicaoAindaValida(paginaEsperada, tokenEsperado)) return;

    const usuarios = jsonUsuarios.usuarios || [];
    const linhas = jsonAuditoria.dados || [];
    if (painel) {
      painel.innerHTML = `
        <section class="exec-panel admin-robo-entry-card-v81729 interface-card-compact-v81746">
          <div class="exec-panel-title">
            <div>
              <h4>Administração do Robô</h4>
              <p>Outlook, launcher e e-mails.</p>
            </div>
            <button class="btn requires-admin" onclick="abrirAdminConfiguracoesRoboV81728()">Abrir</button>
          </div>
        </section>

        <section class="exec-panel">
          <div class="exec-panel-title"><div><h4>Usuários</h4><p>Cadastros, perfis e acessos.</p></div></div>
          <div class="admin-user-form">
            <input id="adminUsuario" placeholder="usuário" />
            <input id="adminNome" placeholder="nome completo" />
            <select id="adminPerfil"><option value="consulta">consulta</option><option value="admin">admin</option></select>
            <input id="adminSenha" type="password" placeholder="senha inicial / nova senha" />
            <label class="admin-check"><input id="adminAtivo" type="checkbox" checked /> Ativo</label>
            <label class="admin-check"><input id="adminTrocarSenha" type="checkbox" checked /> Exigir troca</label>
            <button class="btn" onclick="salvarUsuarioAdmin()">Salvar usuário</button>
          </div>
          <div class="admin-audit-wrap">
            <table class="admin-audit-table">
              <thead><tr><th>Usuário</th><th>Nome</th><th>Perfil</th><th>Status</th><th>Troca senha</th><th>Último login</th><th>Ações</th></tr></thead>
              <tbody>
                ${usuarios.map(u => `
                  <tr>
                    <td>${u.usuario || ""}</td>
                    <td>${u.nome || ""}</td>
                    <td>${u.perfil || ""}</td>
                    <td>${u.ativo ? "Ativo" : "Inativo"}</td>
                    <td>${u.trocar_senha ? "Sim" : "Não"}</td>
                    <td>${u.ultimo_login_em || "-"}</td>
                    <td class="admin-actions-cell">
                      <button class="btn secondary mini" onclick='editarUsuarioAdmin(${JSON.stringify(u)})'>Editar</button>
                      <button class="btn secondary mini" onclick='alternarStatusUsuarioAdmin(${JSON.stringify(u.usuario || "")}, ${u.ativo ? 'false' : 'true'})'>${u.ativo ? "Desativar" : "Ativar"}</button>
                      <button class="btn danger mini" onclick='excluirUsuarioAdmin(${JSON.stringify(u.usuario || "")})'>Excluir</button>
                    </td>
                  </tr>`).join("") || `<tr><td colspan="7">Nenhum usuário cadastrado.</td></tr>`}
              </tbody>
            </table>
          </div>
        </section>

        <section class="exec-panel">
          <div class="exec-panel-title"><div><h4>Minha Senha</h4><p>Alteração do acesso atual.</p></div></div>
          <div class="admin-user-form compact">
            <input id="minhaSenhaAtual" type="password" placeholder="senha atual" />
            <input id="minhaNovaSenha" type="password" placeholder="nova senha" />
            <input id="minhaNovaSenha2" type="password" placeholder="confirmar nova senha" />
            <button class="btn" onclick="alterarMinhaSenha()">Alterar minha senha</button>
          </div>
        </section>

        <section class="exec-panel">
          <div class="exec-panel-title"><div><h4>Auditoria</h4><p>Últimas 200 ações.</p></div></div>
          <div class="admin-audit-wrap">
            <table class="admin-audit-table">
              <thead><tr><th>Quando</th><th>Usuário</th><th>Perfil</th><th>Ação</th><th>Status</th><th>Detalhe</th><th>IP</th></tr></thead>
              <tbody>
                ${linhas.map(l => `<tr><td>${l.quando || ""}</td><td>${l.nome || l.usuario || ""}</td><td>${l.perfil || ""}</td><td>${l.acao || ""}</td><td>${l.status || ""}</td><td>${l.detalhe || ""}</td><td>${l.ip || ""}</td></tr>`).join("") || `<tr><td colspan="7">Nenhum registro de auditoria ainda.</td></tr>`}
              </tbody>
            </table>
          </div>
        </section>
      `;
    }
  } catch (e) {
    if (painel) painel.innerHTML = `<div class="admin-panel"><h3>Falha ao carregar administração</h3><p>${escapeHtml(e.message)}</p></div>`;
  }
}


function payloadConfiguracaoRobo() {
  return {
    arquivo_excel_robo: qs("cfgExcelRobo") ? qs("cfgExcelRobo").value.trim() : "",
    pasta_pdfs_robo: qs("cfgPastaPdfs") ? qs("cfgPastaPdfs").value.trim() : "",
    script_launcher_robo: qs("cfgScriptLauncher") ? qs("cfgScriptLauncher").value.trim() : "",
    max_pdfs_correcao_saldo: qs("cfgMaxPdfs") ? Number(qs("cfgMaxPdfs").value || 1500) : 1500,
    max_segundos_correcao_saldo: qs("cfgMaxSegundos") ? Number(qs("cfgMaxSegundos").value || 180) : 180,
    auto_importar_excel_robo: qs("cfgAutoImportar") ? qs("cfgAutoImportar").checked : true,
    corrigir_saldo_atual_por_pdf_ao_importar: qs("cfgCorrigirPdf") ? qs("cfgCorrigirPdf").checked : false,
  };
}

function renderResultadoConfigRobo(json) {
  const alvo = qs("cfgRoboResultado");
  if (!alvo) return;
  const exemplos = (json.exemplos || []).slice(0, 5).map(x => `<li>${escapeHtml(x)}</li>`).join("");
  const pastas = (json.pastas_verificadas || []).map(x => `<li>${escapeHtml(x)}</li>`).join("");
  alvo.innerHTML = `
    <strong>Excel:</strong> ${json.arquivo_excel_existe || json.arquivo_excel_robo_existe ? "encontrado" : "não encontrado"}<br>
    <strong>Pasta PDFs:</strong> ${json.pasta_pdfs_existe || json.pasta_pdfs_robo_existe ? "encontrada" : "não encontrada"}<br>
    ${json.pdfs_encontrados !== undefined ? `<strong>PDFs encontrados:</strong> ${escapeHtml(String(json.pdfs_encontrados))}<br>` : ""}
    ${pastas ? `<details><summary>Pastas verificadas</summary><ul>${pastas}</ul></details>` : ""}
    ${exemplos ? `<details><summary>Exemplos de PDFs encontrados</summary><ul>${exemplos}</ul></details>` : ""}
  `;
}

async function testarConfiguracaoRobo() {
  try {
    const json = await fetchJsonSeguro("/api/configuracoes-robo/testar", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payloadConfiguracaoRobo()),
    });
    renderResultadoConfigRobo(json);
    alert(`Teste concluído. PDFs encontrados: ${json.pdfs_encontrados ?? 0}`);
  } catch (e) {
    alert("Falha ao testar configuração: " + e.message);
  }
}

async function salvarConfiguracaoRobo() {
  try {
    const json = await fetchJsonSeguro("/api/configuracoes-robo", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payloadConfiguracaoRobo()),
    });
    renderResultadoConfigRobo(json);
    alert(json.mensagem || "Configuração salva com sucesso.");
  } catch (e) {
    alert("Falha ao salvar configuração: " + e.message);
  }
}

async function atualizarHistoricoPdfsBancoHoras() {
  const alvo = qs("cfgRoboResultado");
  if (alvo) alvo.innerHTML = "Atualizando histórico de PDFs. Essa etapa pode demorar, mas não será executada na importação rápida do Excel...";
  try {
    const json = await fetchJsonSeguro("/api/pdfs-banco-horas/atualizar-historico", { method: "POST" });
    if (alvo) alvo.innerHTML = `
      <strong>Histórico de PDFs atualizado</strong><br>
      PDFs encontrados: ${escapeHtml(String(json.pdfs_encontrados ?? 0))}<br>
      Quadros BH lidos: ${escapeHtml(String(json.quadros_lidos ?? 0))}<br>
      Tempo: ${escapeHtml(String(json.segundos ?? ""))}s<br>
      ${escapeHtml(json.mensagem || "")}
    `;
    alert(json.mensagem || "Histórico de PDFs atualizado.");
  } catch (e) {
    if (alvo) alvo.innerHTML = `<span class="erro">Falha ao atualizar histórico: ${escapeHtml(e.message)}</span>`;
    alert("Falha ao atualizar histórico de PDFs: " + e.message);
  }
}

async function corrigirBancoHorasPelosPdfs() {
  const alvo = qs("cfgRoboResultado");
  if (alvo) alvo.innerHTML = "Iniciando correção do Banco de Horas em segundo plano...";
  const payload = { atualizar_cache: false };

  function renderStatus(json, tentativa) {
    const job = json.job || {};
    const resultado = json.resultado || job.resultado || {};
    const msg = json.mensagem || job.mensagem || resultado.mensagem || "";
    if (!alvo) return;

    if (json.em_processamento || json.status === "rodando" || job.status === "rodando") {
      alvo.innerHTML = `
        <strong>Correção do Banco de Horas em andamento</strong><br>
        Processando em segundo plano. Você pode continuar usando o sistema.<br>
        Início: ${escapeHtml(job.iniciado_em || "")}${tentativa ? `<br>Verificação: ${tentativa}` : ""}<br>
        ${escapeHtml(msg)}
      `;
      return;
    }

    if (json.status === "erro" || job.status === "erro" || json.ok === false || resultado.ok === false) {
      alvo.innerHTML = `<span class="erro">Falha ao corrigir Banco de Horas: ${escapeHtml(msg || resultado.erro || "Erro não informado")}</span>`;
      return;
    }

    alvo.innerHTML = `
      <strong>Correção do Banco de Horas concluída</strong><br>
      Linhas corrigidas: ${escapeHtml(String(resultado.linhas_corrigidas ?? json.linhas_corrigidas ?? 0))}<br>
      Linhas sem match: ${escapeHtml(String(resultado.linhas_sem_match ?? json.linhas_sem_match ?? 0))}<br>
      ${resultado.cache_gerado_em || json.cache_gerado_em ? `Cache usado: ${escapeHtml(resultado.cache_gerado_em || json.cache_gerado_em)}<br>` : ""}
      Tempo: ${escapeHtml(String(resultado.segundos ?? job.segundos ?? json.segundos ?? ""))}s<br>
      ${escapeHtml(msg || resultado.mensagem || "")}
    `;
  }

  try {
    let json = await fetchJsonSeguro("/api/banco-horas/corrigir-pelos-pdfs", {
      method: "POST",
      headers: { "Content-Type": "application/json", "Accept": "application/json" },
      credentials: "same-origin",
      body: JSON.stringify(payload),
    });

    renderStatus(json, 0);

    // v8.10.43: a API inicia um job assíncrono; este polling evita o timeout
    // do navegador que gerava "Failed to fetch" enquanto o servidor trabalhava.
    if (json.em_processamento || json.status === "rodando") {
      let finalizado = false;
      for (let i = 1; i <= 60; i++) {
        await new Promise(resolve => setTimeout(resolve, 2500));
        const st = await fetchJsonSeguro("/api/banco-horas/corrigir-pelos-pdfs/status", {
          method: "GET",
          headers: { "Accept": "application/json" },
          credentials: "same-origin",
        });
        renderStatus(st, i);
        if (!st.em_processamento && st.status !== "rodando" && !(st.job && st.job.status === "rodando")) {
          json = st;
          finalizado = true;
          break;
        }
      }
      if (!finalizado && alvo) {
        alvo.innerHTML += `<br><span class="warn">A correção continua em segundo plano. Consulte novamente em alguns minutos ou abra a guia Banco de Horas.</span>`;
      }
    }

    const job = json.job || {};
    const resultado = json.resultado || job.resultado || {};
    if (json.status === "erro" || job.status === "erro" || json.ok === false || resultado.ok === false) {
      alert("Falha ao corrigir Banco de Horas pelos PDFs: " + (json.mensagem || job.mensagem || resultado.mensagem || resultado.erro || "Erro não informado"));
      return;
    }

    alert((json.mensagem || job.mensagem || resultado.mensagem || "Banco de Horas corrigido pelos PDFs."));
    if (paginaAtual === "extrato_banco_horas" && typeof carregarExtratoBancoHoras === "function") carregarExtratoBancoHoras();
    if (paginaAtual === "banco_horas" && typeof carregarBancoHoras === "function") carregarBancoHoras();
  } catch (e) {
    if (alvo) alvo.innerHTML = `<span class="erro">Falha ao corrigir Banco de Horas: ${escapeHtml(e.message || e)}</span>`;
    alert("Falha ao corrigir Banco de Horas pelos PDFs: " + (e.message || e));
  }
}

async function abrirConfiguracaoRobo(btn) {
  iniciarTransicaoPagina("configuracao_robo");
  if (!usuarioAtualEhAdmin()) {
    alert("Acesso negado. Esta área exige perfil Administrador.");
    return;
  }
  paginaAtual = "configuracao_robo";
  document.querySelectorAll(".menu button").forEach(b => b.classList.remove("active"));
  if (btn) btn.classList.add("active");

  qs("tituloPagina").textContent = "Configuração do Robô";
  qs("subtituloPagina").textContent = "Caminhos, limites e importação.";
  marcarPaginaAtualClasse();
  setVisibilidadeDashboardExecutivoPremium(true);
  if (qs("filtros")) qs("filtros").classList.add("hidden");
  if (qs("filtrosDashboard")) qs("filtrosDashboard").classList.add("hidden");
  if (qs("cardsContainer")) qs("cardsContainer").classList.add("hidden");
  if (qs("controleOrdenacaoBancoHoras")) qs("controleOrdenacaoBancoHoras").classList.add("hidden");

  const painel = qs("dashboardExecutivoPremium");
  if (painel) {
    painel.classList.remove("hidden");
    painel.innerHTML = `<div class="admin-panel"><h3>Carregando configuração...</h3></div>`;
  }

  await carregarConfiguracaoRoboDedicada();
}

async function carregarConfiguracaoRoboDedicada() {
  const paginaEsperada = "configuracao_robo";
  const tokenEsperado = tokenAtualPagina();
  const painel = qs("dashboardExecutivoPremium");
  try {
    const [jsonConfigRobo, jsonCachePdfs] = await Promise.all([
      fetchJsonSeguro("/api/configuracoes-robo"),
      fetchJsonSeguro("/api/pdfs-banco-horas/cache"),
    ]);
    const cfgRobo = jsonConfigRobo.config || {};
    const cachePdfs = jsonCachePdfs || {};
    if (painel) {
      painel.innerHTML = `
        <section class="exec-panel interface-card-compact-v81746">
          <div class="admin-config-grid">
            <label>Excel do Robô
              <input id="cfgExcelRobo" value="${escapeHtml(cfgRobo.arquivo_excel_robo || '')}" placeholder="%USERPROFILE%\\ponto_pdfs\\saida\\resultado_lote_consultas_validado.xlsx" />
            </label>
            <label>Pasta dos PDFs
              <input id="cfgPastaPdfs" value="${escapeHtml(cfgRobo.pasta_pdfs_robo || '')}" placeholder="%USERPROFILE%\\ponto_pdfs" />
            </label>
            <label>Script Launcher
              <input id="cfgScriptLauncher" value="${escapeHtml(cfgRobo.script_launcher_robo || '')}" placeholder="Script_Launcher_v5_filtrado.py" />
              <small class="field-help">Opcional quando o script está na pasta do Tempo Fechado.</small>
            </label>
            <label>Limite de PDFs
              <input id="cfgMaxPdfs" type="number" min="10" value="${escapeHtml(String(cfgRobo.max_pdfs_correcao_saldo ?? 1500))}" />
            </label>
            <label>Tempo máximo (segundos)
              <input id="cfgMaxSegundos" type="number" min="10" value="${escapeHtml(String(cfgRobo.max_segundos_correcao_saldo ?? 180))}" />
            </label>
            <label class="admin-check wide"><input id="cfgAutoImportar" type="checkbox" ${cfgRobo.auto_importar_excel_robo === false ? '' : 'checked'} /> Autoimportar Excel</label>
          </div>
          <div class="admin-config-actions">
            <button class="btn secondary" onclick="testarConfiguracaoRobo()">Testar caminhos</button>
            <button class="btn" onclick="salvarConfiguracaoRobo()">Salvar configuração</button>
          </div>
          <div id="cfgRoboResultado" class="admin-config-result">
            Excel existe: ${jsonConfigRobo.arquivo_excel_robo_existe ? 'sim' : 'não'} &nbsp;|&nbsp; Pasta PDFs existe: ${jsonConfigRobo.pasta_pdfs_robo_existe ? 'sim' : 'não'}<br>
            Modelo: Consolidado oficial e histórico de auditoria.
          </div>
          <div class="admin-config-hint">Pastas: ${(jsonConfigRobo.pastas_verificadas || []).map(p => escapeHtml(p)).join(' | ')}</div>
        </section>
      `;
    }
  } catch (e) {
    if (painel) painel.innerHTML = `<div class="admin-panel"><h3>Falha ao carregar configuração</h3><p>${escapeHtml(e.message)}</p></div>`;
  }
}


function editarUsuarioAdmin(u) {
  qs("adminUsuario").value = u.usuario || "";
  qs("adminNome").value = u.nome || "";
  qs("adminPerfil").value = u.perfil || "consulta";
  qs("adminSenha").value = "";
  qs("adminAtivo").checked = !!u.ativo;
  qs("adminTrocarSenha").checked = !!u.trocar_senha;
  qs("adminUsuario").focus();
}

async function salvarUsuarioAdmin() {
  const payload = {
    usuario: qs("adminUsuario").value.trim(),
    nome: qs("adminNome").value.trim(),
    perfil: qs("adminPerfil").value,
    senha: qs("adminSenha").value,
    ativo: qs("adminAtivo").checked,
    trocar_senha: qs("adminTrocarSenha").checked,
  };
  if (!payload.usuario || !payload.nome) { alert("Informe usuário e nome."); return; }
  try {
    const json = await fetchJsonSeguro("/api/usuarios", {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify(payload),
    });
    alert(json.mensagem || "Usuário salvo com sucesso.");
    await carregarAuditoriaMultiusuario();
  } catch (e) {
    alert("Falha ao salvar usuário: " + e.message);
  }
}

async function alternarStatusUsuarioAdmin(usuario, ativo) {
  try {
    await fetchJsonSeguro(`/api/usuarios/${encodeURIComponent(usuario)}/status`, {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify({ativo}),
    });
    await carregarAuditoriaMultiusuario();
  } catch (e) {
    alert("Falha ao alterar status: " + e.message);
  }
}

async function excluirUsuarioAdmin(usuario) {
  usuario = String(usuario || "").trim();
  if (!usuario) return;
  const confirmado = confirm(`Excluir definitivamente o usuario "${usuario}"?\n\nEsta acao remove o cadastro e nao pode ser desfeita.`);
  if (!confirmado) return;
  try {
    const json = await fetchJsonSeguro(`/api/usuarios/${encodeURIComponent(usuario)}`, {
      method: "DELETE",
    });
    alert(json.mensagem || "Usuario excluido com sucesso.");
    await carregarAuditoriaMultiusuario();
  } catch (e) {
    alert("Falha ao excluir usuario: " + e.message);
  }
}

async function alterarMinhaSenha() {
  const senhaAtual = qs("minhaSenhaAtual").value;
  const novaSenha = qs("minhaNovaSenha").value;
  const confirma = qs("minhaNovaSenha2").value;
  if (!novaSenha || novaSenha !== confirma) { alert("A confirmação da nova senha não confere."); return; }
  try {
    const json = await fetchJsonSeguro("/api/minha-senha", {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify({senha_atual: senhaAtual, nova_senha: novaSenha}),
    });
    alert(json.mensagem || "Senha alterada com sucesso.");
    window.USUARIO_ATUAL.trocar_senha = false;
  } catch (e) {
    alert("Falha ao alterar senha: " + e.message);
  }
}

function abrirTrocaSenhaRapida() {
  const senhaAtual = prompt("Informe sua senha atual:");
  if (senhaAtual === null) return;
  const novaSenha = prompt("Informe a nova senha, com pelo menos 6 caracteres:");
  if (novaSenha === null) return;
  const confirma = prompt("Confirme a nova senha:");
  if (confirma === null) return;
  if (novaSenha !== confirma) { alert("A confirmação da nova senha não confere."); return; }
  fetch("/api/minha-senha", {
    method: "POST",
    headers: {"Content-Type": "application/json"},
    body: JSON.stringify({senha_atual: senhaAtual, nova_senha: novaSenha}),
  })
    .then(async resp => {
      const json = await resp.json();
      if (!resp.ok || json.ok === false) throw new Error(json.erro || json.mensagem || `HTTP ${resp.status}`);
      alert(json.mensagem || "Senha alterada com sucesso.");
      if (window.USUARIO_ATUAL) window.USUARIO_ATUAL.trocar_senha = false;
    })
    .catch(e => alert("Falha ao alterar senha: " + e.message));
}

function abrirModalTrocaSenha(obrigatorio=false) {
  if (paginaAtual !== "administracao" && usuarioAtualEhAdmin()) {
    // Para administradores, a troca fica disponível na guia Administração.
    return;
  }
  const aviso = obrigatorio ? "Sua senha é inicial. Altere-a na área Administração > Minha Senha." : "Altere sua senha na área Administração.";
  alert(aviso);
}

function scoreBadgeClasse(c){
  return c || "info";
}
function renderScoreLinha(c){
  const penalidades = (c.penalidades || []).map(p => `<li><span>${escapeHtml(p.motivo || "")}</span><b>${p.pontos || 0}</b><small>${escapeHtml(p.detalhe || "")}</small></li>`).join("");
  return `<div class="score-row ${scoreBadgeClasse(c.classe)}">
    <div class="score-number"><strong>${c.score}</strong><span>${escapeHtml(c.faixa || "")}</span></div>
    <div class="score-main">
      <strong>${escapeHtml(c.nome || "-")}</strong>
      <span>${escapeHtml(c.cc || "-")} • ${escapeHtml(c.turno || "-")} • Banco ${escapeHtml(c.saldo || "00,00")} • H.E. ${escapeHtml(c.he || "00:00")}</span>
      ${penalidades ? `<ul class="score-penalties">${penalidades}</ul>` : `<small>Sem penalidades relevantes no filtro atual.</small>`}
    </div>
  </div>`;
}

function renderScoreGrupo(titulo, linhas){
  linhas = linhas || [];
  return `<section class="exec-panel">
    <div class="exec-panel-title"><div><h4>${titulo}</h4><p>Score médio e quantidade de colaboradores.</p></div></div>
    <div class="score-group-list">
      ${linhas.length ? linhas.map(r => `<div><strong>${escapeHtml(r.grupo || "-")}</strong><span>${r.score_medio}</span><small>${r.colaboradores} colaborador(es)</small></div>`).join("") : `<div class="exec-empty">Sem dados para o filtro.</div>`}
    </div>
  </section>`;
}

function renderScoreOperacional(payload){
  const painel = document.getElementById("dashboardExecutivoPremium");
  if(!painel) return;
  payload = payload || {};
  const k = payload.kpis || {};
  painel.innerHTML = `
    <div class="exec-hero">
      <div>
        <h3>🎯 Score Operacional</h3>
        <p>Nota gerencial de 0 a 100 baseada em banco, H.E., ausências, inconsistências, ponto aberto e jornada.</p>
      </div>
      <div class="exec-periodo"><span>Colaboradores avaliados</span><strong>${k.colaboradores || 0}</strong><small>Score médio: ${k.score_medio || 0}</small></div>
    </div>
    <div class="exec-kpi-grid">
      ${cardResumoExecutivo("Score Médio", k.score_medio ?? 0, "empresa", "")}
      ${cardResumoExecutivo("Excelente", k.excelente ?? 0, "90 a 100", "positive")}
      ${cardResumoExecutivo("Controlado", k.controlado ?? 0, "75 a 89", "")}
      ${cardResumoExecutivo("Atenção", k.atencao ?? 0, "60 a 74", "warning")}
      ${cardResumoExecutivo("Crítico", k.critico ?? 0, "40 a 59", "danger")}
      ${cardResumoExecutivo("Intervenção", k.intervencao ?? 0, "abaixo de 40", "danger")}
    </div>
    <div class="exec-grid exec-grid-2">
      <section class="exec-panel">
        <div class="exec-panel-title"><div><h4>Melhores Scores</h4><p>Colaboradores com maior estabilidade operacional.</p></div></div>
        <div class="score-list">${(payload.melhores || []).map(renderScoreLinha).join("") || `<div class="exec-empty">Sem dados.</div>`}</div>
      </section>
      <section class="exec-panel">
        <div class="exec-panel-title"><div><h4>Piores Scores</h4><p>Prioridade gerencial para investigação e ação.</p></div></div>
        <div class="score-list">${(payload.piores || []).map(renderScoreLinha).join("") || `<div class="exec-empty">Sem dados.</div>`}</div>
      </section>
    </div>
    <div class="exec-grid exec-grid-2">
      ${renderScoreGrupo("Score por Centro de Custo", payload.por_cc)}
      ${renderScoreGrupo("Score por Turno", payload.por_turno)}
    </div>
  `;
}

async function carregarScoreOperacional(){
  const paginaEsperada = "score_operacional";
  const tokenEsperado = tokenAtualPagina();
  const p = new URLSearchParams();
  if(qs("cc") && qs("cc").value) p.set("cc", qs("cc").value);
  if(qs("nome") && qs("nome").value) p.set("nome", qs("nome").value);
  if(qs("turno") && qs("turno").value) p.set("turno", qs("turno").value);
  if(qs("data") && qs("data").value) p.set("data", qs("data").value);
  const { controller, timer } = criarControllerNavegacao(12000);
  try{
    const resp = await fetch(`/api/score-operacional?${p.toString()}`, { signal: controller.signal });
    const json = await resp.json();
    if (!requisicaoAindaValida(paginaEsperada, tokenEsperado)) return;
    if(!resp.ok || json.erro){
      renderScoreOperacional({});
      if(qs("statusCarga")) qs("statusCarga").textContent = json.erro || `Erro HTTP ${resp.status}`;
      return;
    }
    renderScoreOperacional(json);
    if(qs("statusCarga")) qs("statusCarga").textContent = `${json.kpis?.colaboradores || 0} colaborador(es) avaliados no Score Operacional.`;
  }catch(e){
    if (!requisicaoAindaValida(paginaEsperada, tokenEsperado)) return;
    renderScoreOperacional({});
    if(qs("statusCarga")) qs("statusCarga").textContent = e.name === "AbortError" ? "Carregamento do Score cancelado pela troca de guia." : "Erro ao carregar Score Operacional: " + e.message;
  } finally {
    finalizarControllerNavegacao(controller, timer);
  }
}

function abrirScoreOperacional(el){
  iniciarTransicaoPagina("score_operacional");
  paginaAtual="score_operacional";
  marcarPaginaAtualClasse();
  setVisibilidadeDashboardExecutivoPremium(true);
  dadosConsolidadoAtual = [];
  document.querySelectorAll(".menu button").forEach(b=>b.classList.remove("active"));
  if(el) el.classList.add("active");
  if(qs("tituloPagina")) qs("tituloPagina").textContent="Score Operacional";
  if(qs("subtituloPagina")) qs("subtituloPagina").textContent="";
  if(qs("tituloTabela")) qs("tituloTabela").textContent="Score Operacional";
  if(qs("statusCarga")) qs("statusCarga").textContent="Carregando Score Operacional...";
  if(qs("filtros")) qs("filtros").classList.remove("hidden");
  if(qs("filtrosDashboard")) qs("filtrosDashboard").classList.add("hidden");
  if(qs("boxFiltroTipo")) qs("boxFiltroTipo").classList.add("hidden");
  if(qs("controleOrdenacaoBancoHoras")) qs("controleOrdenacaoBancoHoras").classList.add("hidden");
  if(qs("tabelaDashboard")) qs("tabelaDashboard").classList.add("hidden");
  if(qs("tabelaConsolidado")) qs("tabelaConsolidado").classList.add("hidden");
  ocultarFiltroDataPrincipalPremium();
  carregarScoreOperacional();
}


async function abrirConfigNotificacoesJornada() {
  if (!usuarioAtualEhAdmin()) {
    alert("Acesso negado. A configuração de notificações exige perfil Administrador.");
    return;
  }
  try {
    const resp = await fetch("/api/notificacoes-jornada/config");
    const json = await resp.json();
    if (!resp.ok || json.ok === false) throw new Error(json.erro || `Erro HTTP ${resp.status}`);
    const c = json.config || {};
    const html = `
Configuração de Notificações de Jornada

E-mail ativo? ${c.EMAIL_ATIVO ? "sim" : "não"}
Digite os dados no próximo formulário simplificado.
Arquivo: ${c._arquivo_config || ""}
`;
    const emailAtivo = confirm("Ativar notificações por e-mail?");
    const destinatarios = prompt("Destinatários de e-mail separados por vírgula:", (c.EMAIL_DESTINATARIOS || []).join(", "));
    const smtpHost = prompt("SMTP_HOST:", c.SMTP_HOST || "");
    const smtpPort = prompt("SMTP_PORT:", c.SMTP_PORT || 587);
    const smtpUsuario = prompt("SMTP_USUARIO:", c.SMTP_USUARIO || "");
    const smtpSenha = prompt("SMTP_SENHA (deixe ******** para manter):", c.SMTP_SENHA || "");
    const remetente = prompt("EMAIL_REMETENTE:", c.EMAIL_REMETENTE || c.SMTP_USUARIO || "");
    const enviarAoImportar = confirm("Enviar automaticamente ao importar Excel quando houver novas ocorrências?");

    const payload = {
      EMAIL_ATIVO: emailAtivo,
      EMAIL_DESTINATARIOS: destinatarios || "",
      SMTP_HOST: smtpHost || "",
      SMTP_PORT: Number(smtpPort || 587),
      SMTP_USAR_TLS: true,
      SMTP_USUARIO: smtpUsuario || "",
      SMTP_SENHA: smtpSenha || "",
      EMAIL_REMETENTE: remetente || "",
      ENVIAR_AO_IMPORTAR_EXCEL: enviarAoImportar,
    };

    const save = await fetch("/api/notificacoes-jornada/config", {
      method: "POST",
      headers: {"Content-Type": "application/json", "Accept": "application/json"},
      body: JSON.stringify(payload)
    });
    const saved = await save.json();
    if (!save.ok || saved.ok === false) throw new Error(saved.erro || `Erro HTTP ${save.status}`);
    alert(saved.mensagem || "Configuração salva.");
    carregarAlertasAutomaticos();
  } catch (e) {
    alert("Erro ao configurar notificações: " + e.message);
  }
}


async function abrirConfigEmailRoboV81727() {
  if (!usuarioAtualEhAdmin()) {
    alert("Acesso negado. A configuração do e-mail do robô exige perfil Administrador.");
    return;
  }
  try {
    const resp = await fetch("/api/config-email-robo");
    const json = await resp.json();
    if (!resp.ok || json.ok === false) throw new Error(json.erro || `Erro HTTP ${resp.status}`);
    const c = json.config || {};
    const enviar = confirm("Enviar automaticamente após processar?");
    const revisar = confirm("Revisar no Outlook antes do envio?");
    const destinatarios = prompt("Para separados por vírgula:", (c.destinatarios_email || []).join(", "));
    if (destinatarios === null) return;
    const cc = prompt("CC separados por vírgula:", (c.cc_email || []).join(", "));
    if (cc === null) return;
    const cco = prompt("CCO separados por vírgula:", (c.cco_email || []).join(", "));
    if (cco === null) return;
    const assunto = prompt("Assunto do e-mail:", c.assunto_email_base || "Arquivo consolidado");
    if (assunto === null) return;
    const payload = {
      enviar_email_automaticamente: enviar,
      exibir_email_antes_de_enviar: revisar,
      destinatarios_email: destinatarios,
      cc_email: cc,
      cco_email: cco,
      assunto_email_base: assunto
    };
    const save = await fetch("/api/config-email-robo", {
      method: "POST",
      headers: {"Content-Type": "application/json", "Accept": "application/json"},
      body: JSON.stringify(payload)
    });
    const saved = await save.json();
    if (!save.ok || saved.ok === false) throw new Error(saved.erro || `Erro HTTP ${save.status}`);
    alert((saved.mensagem || "Configuração salva.") + "\n\nArquivo:\n" + (saved.arquivo || ""));
  } catch (e) {
    alert("Erro ao configurar e-mail do robô: " + e.message);
  }
}

function adicionarBotaoConfigEmailRoboV81727() {
  try {
    const alvo = qs("dashboardExecutivoPremium") || qs("conteudoPrincipal") || document.body;
    if (!alvo || document.getElementById("btnConfigEmailRoboV81727")) return;
    const wrap = document.createElement("div");
    wrap.id = "btnConfigEmailRoboV81727";
    wrap.style.margin = "10px 0";
    wrap.innerHTML = `<button class="btn requires-admin" onclick="abrirConfigEmailRoboV81727()">Configurar e-mail do robô</button>`;
    alvo.prepend(wrap);
  } catch(e) {}
}
const _abrirPaginaOriginalV81727 = typeof abrirPagina === "function" ? abrirPagina : null;
if (_abrirPaginaOriginalV81727) {
  abrirPagina = function(pagina, el) {
    const r = _abrirPaginaOriginalV81727(pagina, el);
    if (pagina === "processamento" || pagina === "processamento_integrado") {
      setTimeout(adicionarBotaoConfigEmailRoboV81727, 300);
    }
    return r;
  }
}


function campoAdminRoboV81728(id, label, valor, tipo = "text", dica = "") {
  return `<label class="admin-robo-field"><span>${label}</span><input id="${id}" type="${tipo}" value="${escapeHtml(valor || "")}" />${dica ? `<small>${dica}</small>` : ""}</label>`;
}
function checkboxAdminRoboV81728(id, label, marcado, dica = "") {
  return `<label class="admin-robo-check"><input id="${id}" type="checkbox" ${marcado ? "checked" : ""} /><span>${label}</span>${dica ? `<small>${dica}</small>` : ""}</label>`;
}
function renderAdminConfiguracoesRoboV81728(painel) {
  painel = painel || {};
  const email = painel.email_robo || {};
  const launcher = painel.launcher || {};
  const outlook = painel.outlook || {};
  const arquivos = painel.arquivos || {};
  const alvo = qs("dashboardExecutivoPremium") || qs("conteudoPrincipal") || document.body;
  if (!alvo) return;
  alvo.innerHTML = `
    <div class="admin-robo-shell">
      <div class="admin-robo-actions-top">
        <button class="btn requires-admin" onclick="salvarAdminConfiguracoesRoboV81728()">Salvar</button>
      </div>
      <div class="admin-robo-layout">
        <aside class="admin-robo-tree">
          <button class="active" onclick="selecionarAdminRoboSecaoV81728('launcher', this)">Launcher</button>
          <button onclick="selecionarAdminRoboSecaoV81728('outlook', this)">Outlook</button>
          <button onclick="selecionarAdminRoboSecaoV81728('email', this)">E-mail</button>
        </aside>
        <section class="admin-robo-content">
          <div class="admin-robo-section" data-admin-robo-section="launcher">
            <h4>Launcher</h4>
            <p>Script usado no processamento integrado.</p>
            ${campoAdminRoboV81728("adminRoboLauncherPath", "Script Launcher", launcher.script_launcher_robo || "", "text", "Informativo nesta versão.")}
            <div class="admin-robo-note">Arquivo de configuração: ${escapeHtml(launcher.arquivo_config_robo || "")}</div>
          </div>
          <div class="admin-robo-section hidden" data-admin-robo-section="outlook">
            <h4>Outlook</h4>
            <p>Filtros para localizar PDFs por e-mail.</p>
            ${campoAdminRoboV81728("adminRoboOutlookRemetente", "Remetente", outlook.remetente_filtro || "", "text", "Filtro de origem.")}
            ${campoAdminRoboV81728("adminRoboOutlookAssunto", "Assunto", outlook.assunto_filtro || "", "text", "Filtro do título do e-mail.")}
            ${campoAdminRoboV81728("adminRoboLimiteEmailsOutlook", "E-mails recentes a analisar", String(outlook.limite_emails_outlook ?? 500), "number", `De ${outlook.limite_emails_outlook_min || 50} a ${outlook.limite_emails_outlook_max || 5000}. Use valores maiores quando houver muitos e-mails na caixa.`)}
            ${checkboxAdminRoboV81728("adminRoboCapturarEmailsLidos", "Incluir e-mails lidos recentes", !!outlook.capturar_emails_lidos, "Desmarcado: somente não lidos. Marcado: inclui lidos recentes.")}
            <div class="admin-robo-note"><strong>Política atual:</strong> ${escapeHtml(outlook.politica_captura || "Somente e-mails não lidos")} · <strong>Janela:</strong> ${escapeHtml(String(outlook.limite_emails_outlook ?? 500))} e-mail(s)</div>
            <div class="admin-robo-note">${escapeHtml(outlook.observacao || "")}</div>
          </div>
          <div class="admin-robo-section hidden" data-admin-robo-section="email">
            <h4>E-mail</h4>
            <p>Destinatários do consolidado.</p>
            ${checkboxAdminRoboV81728("adminRoboEnviarEmail", "Enviar automaticamente após processar", !!email.enviar_email_automaticamente)}
            ${checkboxAdminRoboV81728("adminRoboRevisarOutlook", "Revisar no Outlook antes do envio", !!email.exibir_email_antes_de_enviar)}
            ${campoAdminRoboV81728("adminRoboDestinatarios", "Para", (email.destinatarios_email || []).join(", "), "text", "Separe por vírgula.")}
            ${campoAdminRoboV81728("adminRoboCc", "CC", (email.cc_email || []).join(", "), "text", "Opcional.")}
            ${campoAdminRoboV81728("adminRoboCco", "CCO", (email.cco_email || []).join(", "), "text", "Opcional.")}
            ${campoAdminRoboV81728("adminRoboAssunto", "Assunto", email.assunto_email_base || "Arquivo consolidado")}
            <div class="admin-robo-note">Arquivo salvo em: ${escapeHtml(arquivos.config_email_robo || "")}</div>
          </div>
        </section>
      </div>
    </div>`;
}
function selecionarAdminRoboSecaoV81728(secao, botao) {
  document.querySelectorAll(".admin-robo-tree button").forEach(b => b.classList.remove("active"));
  if (botao) botao.classList.add("active");
  document.querySelectorAll("[data-admin-robo-section]").forEach(el => {
    el.classList.toggle("hidden", el.getAttribute("data-admin-robo-section") !== secao);
  });
}
async function abrirAdminConfiguracoesRoboV81728() {
  iniciarTransicaoPagina("admin_config_robo");
  paginaAtual = "admin_config_robo";
  document.querySelectorAll(".menu button").forEach(b => b.classList.remove("active"));
  if (!usuarioAtualEhAdmin()) {
    alert("Acesso negado. A configuração do robô exige perfil Administrador.");
    return;
  }
  try {
    setVisibilidadeDashboardExecutivoPremium(true);
    if (qs("tituloPagina")) qs("tituloPagina").textContent = "Administração do Robô";
    if (qs("subtituloPagina")) qs("subtituloPagina").textContent = "";
    if (qs("statusCarga")) qs("statusCarga").textContent = "Carregando configurações do robô...";
    const resp = await fetch("/api/admin/config-robo-painel");
    const json = await resp.json();
    if (!resp.ok || json.ok === false) throw new Error(json.erro || `Erro HTTP ${resp.status}`);
    renderAdminConfiguracoesRoboV81728(json.painel || {});
    if (qs("statusCarga")) qs("statusCarga").textContent = "Configurações carregadas.";
  } catch (e) {
    if (qs("statusCarga")) qs("statusCarga").textContent = "Erro ao carregar configurações do robô: " + e.message;
    alert("Erro ao carregar configurações do robô: " + e.message);
  }
}
async function salvarAdminConfiguracoesRoboV81728() {
  if (!usuarioAtualEhAdmin()) { alert("Acesso negado."); return; }
  const payload = {
    robo: {
      remetente_filtro: document.getElementById("adminRoboOutlookRemetente")?.value || "",
      assunto_filtro: document.getElementById("adminRoboOutlookAssunto")?.value || "",
      capturar_emails_lidos: !!document.getElementById("adminRoboCapturarEmailsLidos")?.checked,
      limite_emails_outlook: Number(document.getElementById("adminRoboLimiteEmailsOutlook")?.value || 500)
    },
    email_robo: {
      enviar_email_automaticamente: !!document.getElementById("adminRoboEnviarEmail")?.checked,
      exibir_email_antes_de_enviar: !!document.getElementById("adminRoboRevisarOutlook")?.checked,
      destinatarios_email: document.getElementById("adminRoboDestinatarios")?.value || "",
      cc_email: document.getElementById("adminRoboCc")?.value || "",
      cco_email: document.getElementById("adminRoboCco")?.value || "",
      assunto_email_base: document.getElementById("adminRoboAssunto")?.value || "Arquivo consolidado"
    }
  };
  try {
    const resp = await fetch("/api/admin/config-robo-painel", {
      method: "POST",
      headers: {"Content-Type": "application/json", "Accept": "application/json"},
      body: JSON.stringify(payload)
    });
    const json = await resp.json();
    if (!resp.ok || json.ok === false) throw new Error(json.erro || `Erro HTTP ${resp.status}`);
    alert((json.mensagem || "Configurações salvas.") + "\n\nArquivo e-mail:\n" + (json.arquivo || "") + "\n\nArquivo robô:\n" + (json.arquivo_config_robo || ""));
    abrirAdminConfiguracoesRoboV81728();
  } catch (e) {
    alert("Erro ao salvar configurações do robô: " + e.message);
  }
}

// v8.17.33: injeção dinâmica removida; entrada Administração do Robô fica fixa no HTML.


async function abrirProcessamentoManualRoboV81737() {
  if (!usuarioAtualEhAdmin()) { alert("Acesso negado. O processamento manual do robô exige perfil Administrador."); return; }
  marcarGuiaAtivaV81738("manual");
  limparEstadoOperacionalV81738();
  try {
    const tokenEsperado = iniciarTransicaoPagina("processamento_manual_robo");
    paginaAtual = "processamento_manual_robo";
    marcarPaginaAtualClasse();
    setVisibilidadeDashboardExecutivoPremium(true);
    if (timerProcessamentoIntegrado) { clearInterval(timerProcessamentoIntegrado); timerProcessamentoIntegrado = null; }
    if (qs("filtros")) qs("filtros").classList.add("hidden");
    if (qs("filtrosDashboard")) qs("filtrosDashboard").classList.add("hidden");
    if (qs("cardsContainer")) qs("cardsContainer").classList.add("hidden");
    if (qs("controleOrdenacaoBancoHoras")) qs("controleOrdenacaoBancoHoras").classList.add("hidden");
    if (qs("tituloPagina")) qs("tituloPagina").textContent = "Processamento Manual";
    if (qs("subtituloPagina")) qs("subtituloPagina").textContent = "";
    if (qs("statusCarga")) qs("statusCarga").textContent = "Carregando módulo manual...";
    const resp = await fetch("/api/processamento-manual-robo/status");
    const json = await resp.json();
    if (!requisicaoAindaValida("processamento_manual_robo", tokenEsperado)) return;
    if (!resp.ok || json.ok === false) throw new Error(json.erro || `Erro HTTP ${resp.status}`);
    renderProcessamentoManualRoboV81737(json);
    if (qs("statusCarga")) qs("statusCarga").textContent = "Módulo manual pronto.";
  } catch (e) { alert("Erro ao carregar processamento manual: " + e.message); }
}
function renderProcessamentoManualRoboV81737(status) {
  const alvo = qs("dashboardExecutivoPremium") || qs("conteudoPrincipal") || document.body;
  const pdfs = status.pdfs_entrada || [];
  alvo.innerHTML = `
    <div class="manual-robo-shell">
      <div class="manual-robo-hero">
        <div><h3>Execução do Robô</h3><p>Envie os PDFs pela plataforma ou use arquivos já disponíveis em <strong>ponto_pdfs\\entrada</strong>.</p></div>
        <button class="btn requires-admin" onclick="executarRoboManualV81737()">Executar robô e importar Excel</button>
      </div>
      <div class="manual-robo-grid">
        <section class="exec-panel"><div class="exec-panel-title"><div><h4>Upload de PDFs</h4><p>Carregue os arquivos para a entrada do robô.</p></div></div><input id="manualRoboPdfUpload" type="file" accept="application/pdf,.pdf" multiple /><div class="manual-robo-actions"><button class="btn" type="button" onclick="uploadPdfsRoboManualV82138()">Enviar PDFs</button></div><div id="manualRoboUploadStatus" class="admin-robo-note"></div></section>
        <section class="exec-panel"><div class="exec-panel-title"><div><h4>Pastas e arquivos</h4><p>Estrutura operacional do robô.</p></div></div><div class="admin-robo-note"><strong>Entrada:</strong> ${escapeHtml(status.entrada || "")}</div><div class="admin-robo-note"><strong>Saída:</strong> ${escapeHtml(status.saida || "")}</div><div class="admin-robo-note"><strong>Script:</strong> ${escapeHtml(status.script_robo || "")}</div><div class="admin-robo-note"><strong>Excel padrão:</strong> ${escapeHtml(status.excel_gerado || "")}</div></section>
        <section class="exec-panel interface-card-compact-v81746"><div class="exec-panel-title"><div><h4>Status</h4></div></div><div class="kpi-grid"><div class="kpi-card"><span>Arquivos de entrada</span><strong>${status.qtd_pdfs_entrada || 0}</strong></div><div class="kpi-card"><span>Script localizado</span><strong>${status.script_existe ? "Sim" : "Não"}</strong></div><div class="kpi-card"><span>Excel gerado</span><strong>${status.excel_existe ? "Sim" : "Não"}</strong></div></div><div class="manual-robo-actions"><button type="button" class="btn" onclick="abrirPastaEntradaRoboV81740()">Abrir entrada</button><button class="btn" onclick="importarExcelRoboManualV81737()">Importar Excel</button><button class="btn" onclick="verLogRoboManualV81737()">Ver log</button><button class="btn" onclick="abrirProcessamentoManualRoboV81737()">Recarregar</button></div></section>
      </div>
      <section class="exec-panel"><div class="exec-panel-title"><div><h4>PDFs encontrados na entrada</h4><p>Lista limitada aos 50 primeiros arquivos.</p></div></div><div class="manual-robo-list">${pdfs.length ? pdfs.map(p => `<div>${escapeHtml(p)}</div>`).join("") : "<em>Nenhum PDF encontrado na pasta de entrada.</em>"}</div></section>
    </div>`;
}
async function uploadPdfsRoboManualV82138() {
  const input = document.getElementById("manualRoboPdfUpload");
  const status = document.getElementById("manualRoboUploadStatus");
  const arquivos = input && input.files ? Array.from(input.files) : [];
  if (!arquivos.length) { alert("Selecione ao menos um PDF."); return; }
  const form = new FormData();
  arquivos.forEach(arquivo => form.append("pdfs", arquivo));
  try {
    if (status) status.textContent = "Enviando PDFs...";
    const resp = await fetch("/api/processamento-manual-robo/upload-pdfs", { method: "POST", body: form });
    const json = await resp.json().catch(() => ({}));
    if (!resp.ok || json.ok === false) throw new Error(json.erro || `Erro HTTP ${resp.status}`);
    const ignorados = json.qtd_ignorados ? ` Ignorados: ${json.qtd_ignorados}.` : "";
    if (status) status.textContent = `${json.qtd_salvos || 0} PDF(s) enviados.${ignorados}`;
    alert(json.mensagem || "PDFs enviados para a entrada do robô.");
    abrirProcessamentoManualRoboV81737();
  } catch (e) {
    if (status) status.textContent = "Falha no upload.";
    alert("Erro ao enviar PDFs: " + (e && e.message ? e.message : e));
  }
}
async function abrirPastaEntradaRoboV81740() {
  try {
    const resp = await fetch("/api/processamento-manual-robo/abrir-entrada", {method: "POST"});
    const json = await resp.json().catch(() => ({}));
    if (!resp.ok || json.ok === false) throw new Error(json.erro || `Erro HTTP ${resp.status}`);
    if (json.mensagem) {
      const detalhe = json.entrada ? `\n\n${json.entrada}` : "";
      alert(json.mensagem + detalhe);
    }
  } catch (e) {
    alert("Erro ao abrir pasta de entrada: " + (e && e.message ? e.message : e));
  }
}
window.abrirPastaEntradaRoboV81740 = abrirPastaEntradaRoboV81740;
window.abrirPastaEntradaRoboV81737 = abrirPastaEntradaRoboV81740;
async function executarRoboManualV81737() {
  if (!confirm("Executar o robô sobre os PDFs da entrada e importar automaticamente o Excel padrão?")) return;
  try {
    if (qs("statusCarga")) qs("statusCarga").textContent = "Executando robô manual e importando Excel...";
    const resp = await fetch("/api/processamento-manual-robo/executar", {method: "POST"});
    const json = await resp.json();
    if (!resp.ok || json.ok === false) throw new Error(json.erro || `Erro HTTP ${resp.status}`);
    const imp = json.importacao || {};
    alert((json.mensagem || "Robô executado.") + "\\n\\nExcel: " + (json.excel || "") + "\\n\\nImportação: " + (imp.mensagem || ""));
    abrirProcessamentoManualRoboV81737();
  } catch (e) { alert("Erro ao executar robô manual: " + e.message); verLogRoboManualV81737(); }
}
async function importarExcelRoboManualV81737() {
  try {
    const resp = await fetch("/api/processamento-manual-robo/importar-excel", {method: "POST"});
    const json = await resp.json();
    if (!resp.ok || json.ok === false) throw new Error(json.erro || `Erro HTTP ${resp.status}`);
    const imp = json.importacao || {};
    alert((json.mensagem || imp.mensagem || "Importação concluída.") + "\\n\\nExcel: " + (json.excel || ""));
    abrirProcessamentoManualRoboV81737();
  } catch (e) { alert("Erro ao importar Excel padrão: " + e.message); }
}
async function verLogRoboManualV81737() {
  try {
    const resp = await fetch("/api/processamento-manual-robo/log");
    const json = await resp.json();
    if (!resp.ok || json.ok === false) throw new Error(json.erro || `Erro HTTP ${resp.status}`);
    const alvo = qs("dashboardExecutivoPremium") || qs("conteudoPrincipal") || document.body;
    const antigo = document.getElementById("manualRoboLogPanel");
    if (antigo) antigo.remove();
    alvo.innerHTML += `<section id="manualRoboLogPanel" class="exec-panel"><div class="exec-panel-title"><div><h4>Log técnico do processamento manual</h4></div></div><pre class="manual-robo-log">${escapeHtml(json.log || "")}</pre></section>`;
  } catch (e) { alert("Erro ao carregar log: " + e.message); }
}





// ============================================================
// v8.21.38 - Painel de Implantação
// ============================================================
let painelImplantacaoCacheV8210 = [];

function abrirPainelImplantacaoV8210(el) {
  iniciarTransicaoPagina("painel_implantacao");
  paginaAtual = "painel_implantacao";
  marcarPaginaAtualClasse();
  setVisibilidadeDashboardExecutivoPremium(true);
  dadosConsolidadoAtual = [];

  document.querySelectorAll(".menu button").forEach(b => b.classList.remove("active"));
  if (el) el.classList.add("active");

  if (qs("tituloPagina")) qs("tituloPagina").textContent = "Painel de Implantação";
  if (qs("subtituloPagina")) qs("subtituloPagina").textContent = "Controle dos notebooks instalados, homologados e pendentes.";
  if (qs("tituloTabela")) qs("tituloTabela").textContent = "Painel de Implantação";
  if (qs("statusCarga")) qs("statusCarga").textContent = "Carregando Painel de Implantação...";

  if (qs("filtros")) qs("filtros").classList.add("hidden");
  if (qs("filtrosDashboard")) qs("filtrosDashboard").classList.add("hidden");
  if (qs("boxFiltroTipo")) qs("boxFiltroTipo").classList.add("hidden");
  if (qs("controleOrdenacaoBancoHoras")) qs("controleOrdenacaoBancoHoras").classList.add("hidden");
  if (qs("tabelaDashboard")) qs("tabelaDashboard").classList.add("hidden");
  if (qs("tabelaConsolidado")) qs("tabelaConsolidado").classList.add("hidden");
  ocultarFiltroDataPrincipalPremium();

  carregarPainelImplantacaoV8210();
}

function nomeStatusImplantacaoV8210(status) {
  const mapa = {
    pendente: "Pendente",
    em_configuracao: "Em configuração",
    homologado: "Homologado",
    atencao: "Atenção",
    erro: "Erro",
    desativado: "Desativado"
  };
  return mapa[status] || "Pendente";
}

function classeStatusImplantacaoV8210(status) {
  if (status === "homologado") return "ok";
  if (status === "atencao" || status === "em_configuracao") return "warning";
  if (status === "erro") return "danger";
  if (status === "desativado") return "muted";
  return "info";
}

function checksConcluidosImplantacaoV8210(n) {
  const c = n.checklist || {};
  const vals = Object.values(c);
  const total = vals.length || 8;
  const ok = vals.filter(Boolean).length;
  return { ok, total, pct: Math.round((ok / total) * 100) };
}

function renderNotebookImplantacaoV8210(n) {
  const status = n.status || "pendente";
  const cls = classeStatusImplantacaoV8210(status);
  const checks = checksConcluidosImplantacaoV8210(n);
  return `
    <div class="implant-card ${cls}">
      <div class="implant-card-head">
        <div>
          <strong>${valorSeguro(n.nome, "Notebook")}</strong>
          <span>${valorSeguro(n.responsavel, "Sem responsável")} · ${valorSeguro(n.setor, "Sem setor")}</span>
        </div>
        <b class="implant-status ${cls}">${nomeStatusImplantacaoV8210(status)}</b>
      </div>
      <div class="implant-meta">
        <span>Versão: <b>${valorSeguro(n.versao, "-")}</b></span>
        <span>Homologação: <b>${valorSeguro(n.ultima_homologacao, "-")}</b></span>
        <span>Checklist: <b>${checks.ok}/${checks.total}</b></span>
      </div>
      <div class="implant-progress"><span style="width:${checks.pct}%"></span></div>
      <p>${valorSeguro(n.observacoes, "Sem observações registradas.")}</p>
      <div class="implant-card-actions">
        <button class="btn secondary" onclick="editarNotebookImplantacaoV8210('${n.id}')">Editar</button>
        <button class="btn secondary" onclick="excluirNotebookImplantacaoV8210('${n.id}')">Remover</button>
      </div>
    </div>
  `;
}

function renderFormularioImplantacaoV8210(n = {}) {
  const c = n.checklist || {};
  const id = valorSeguro(n.id, "");
  return `
    <section class="exec-panel implant-form-panel">
      <div class="exec-panel-title">
        <div>
          <h4>${id ? "Editar notebook" : "Adicionar notebook"}</h4>
          <p>Registro local para acompanhar instalação, configuração e homologação.</p>
        </div>
        <button class="btn secondary" onclick="limparFormularioImplantacaoV8210()">Limpar</button>
      </div>
      <form id="formPainelImplantacaoV8210" class="implant-form" onsubmit="salvarNotebookImplantacaoV8210(event)">
        <input type="hidden" name="id" value="${id}">
        <label>Notebook<input name="nome" value="${valorSeguro(n.nome, "")}" placeholder="Ex.: Notebook Financeiro 01"></label>
        <label>Responsável<input name="responsavel" value="${valorSeguro(n.responsavel, "")}" placeholder="Nome do usuário/responsável"></label>
        <label>Setor<input name="setor" value="${valorSeguro(n.setor, "")}" placeholder="Ex.: RH, DP, Operação"></label>
        <label>Status
          <select name="status">
            ${["pendente","em_configuracao","homologado","atencao","erro","desativado"].map(st => `<option value="${st}" ${st === (n.status || "pendente") ? "selected" : ""}>${nomeStatusImplantacaoV8210(st)}</option>`).join("")}
          </select>
        </label>
        <label>Versão<input name="versao" value="${valorSeguro(n.versao, VERSAO_ESPERADA_BADGE)}" placeholder="versao atual"></label>
        <label>Última homologação<input name="ultima_homologacao" value="${valorSeguro(n.ultima_homologacao, "")}" placeholder="AAAA-MM-DD"></label>
        <label class="implant-wide">Observações<textarea name="observacoes" rows="3" placeholder="Pendências, exceções, caminho do Excel, observações do notebook...">${valorSeguro(n.observacoes, "")}</textarea></label>
        <div class="implant-checklist implant-wide">
          ${[
            ["instalador", "Instalador aplicado"],
            ["login", "Login testado"],
            ["outlook", "Outlook configurado"],
            ["pastas", "Pastas validadas"],
            ["banco_horas", "Banco de Horas OK"],
            ["central_operacao", "Central de Operação OK"],
            ["processamento", "Processamento testado"],
            ["backup_seguro", "Backup seguro OK"]
          ].map(([key, label]) => `<label><input type="checkbox" name="check_${key}" ${c[key] ? "checked" : ""}> ${label}</label>`).join("")}
        </div>
        <div class="implant-wide implant-actions">
          <button class="btn" type="submit">Salvar notebook</button>
        </div>
      </form>
    </section>
  `;
}

function dadosFormularioImplantacaoV8210() {
  const form = qs("formPainelImplantacaoV8210");
  const fd = new FormData(form);
  return {
    id: fd.get("id") || "",
    nome: fd.get("nome") || "Notebook",
    responsavel: fd.get("responsavel") || "",
    setor: fd.get("setor") || "",
    status: fd.get("status") || "pendente",
    versao: fd.get("versao") || VERSAO_ESPERADA_BADGE,
    ultima_homologacao: fd.get("ultima_homologacao") || "",
    observacoes: fd.get("observacoes") || "",
    checklist: {
      instalador: !!fd.get("check_instalador"),
      login: !!fd.get("check_login"),
      outlook: !!fd.get("check_outlook"),
      pastas: !!fd.get("check_pastas"),
      banco_horas: !!fd.get("check_banco_horas"),
      central_operacao: !!fd.get("check_central_operacao"),
      processamento: !!fd.get("check_processamento"),
      backup_seguro: !!fd.get("check_backup_seguro")
    }
  };
}

async function salvarNotebookImplantacaoV8210(ev) {
  ev.preventDefault();
  try {
    const resp = await fetch("/api/painel-implantacao/notebooks", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(dadosFormularioImplantacaoV8210())
    });
    const json = await resp.json();
    if (!resp.ok || json.ok === false) throw new Error(json.erro || `Erro HTTP ${resp.status}`);
    carregarPainelImplantacaoV8210();
  } catch (e) {
    alert("Erro ao salvar notebook: " + e.message);
  }
}

async function excluirNotebookImplantacaoV8210(id) {
  if (!confirm("Remover este notebook do Painel de Implantação?")) return;
  try {
    const resp = await fetch(`/api/painel-implantacao/notebooks/${encodeURIComponent(id)}`, { method: "DELETE" });
    const json = await resp.json();
    if (!resp.ok || json.ok === false) throw new Error(json.erro || `Erro HTTP ${resp.status}`);
    carregarPainelImplantacaoV8210();
  } catch (e) {
    alert("Erro ao remover notebook: " + e.message);
  }
}

function editarNotebookImplantacaoV8210(id) {
  const n = painelImplantacaoCacheV8210.find(x => String(x.id) === String(id));
  const formHost = qs("implantacaoFormHostV8210");
  if (formHost) formHost.innerHTML = renderFormularioImplantacaoV8210(n || {});
  if (formHost) formHost.scrollIntoView({ behavior: "smooth", block: "start" });
}

function limparFormularioImplantacaoV8210() {
  const formHost = qs("implantacaoFormHostV8210");
  if (formHost) formHost.innerHTML = renderFormularioImplantacaoV8210({ versao: VERSAO_ESPERADA_BADGE });
}

function preencherNotebookLocalImplantacaoV8210(local) {
  const hoje = new Date().toISOString().slice(0, 10);
  const n = {
    nome: local?.hostname || "Notebook local",
    responsavel: local?.usuario_windows || "",
    setor: "",
    status: "em_configuracao",
    versao: VERSAO_ESPERADA_BADGE,
    ultima_homologacao: hoje,
    observacoes: `Perfil local: ${local?.perfil || "-"}\nSistema: ${local?.sistema || "-"}`,
    checklist: { instalador: true, login: false, outlook: false, pastas: false, banco_horas: false, central_operacao: true, processamento: false, backup_seguro: true }
  };
  const formHost = qs("implantacaoFormHostV8210");
  if (formHost) formHost.innerHTML = renderFormularioImplantacaoV8210(n);
}

function renderPainelImplantacaoV8210(payload) {
  const painel = qs("dashboardExecutivoPremium");
  if (!painel) return;
  const r = payload.resumo || {};
  const notebooks = payload.notebooks || [];
  painelImplantacaoCacheV8210 = notebooks;
  painel.innerHTML = `
    <div class="op-hero implant-hero">
      <div>
        <h3>Painel de Implantação</h3>
      </div>
      <div class="implant-hero-actions">
        <button class="btn secondary" onclick='preencherNotebookLocalImplantacaoV8210(${JSON.stringify(payload.notebook_local || {}).replace(/'/g, "&#39;")})'>Usar notebook local</button>
        <button class="btn" onclick="carregarPainelImplantacaoV8210()">Atualizar</button>
      </div>
    </div>

    <div class="exec-kpi-grid">
      ${cardResumoExecutivo("Notebooks", r.total ?? 0, "cadastrados", "")}
      ${cardResumoExecutivo("Homologados", r.homologados ?? 0, "prontos", (r.homologados || 0) > 0 ? "positive" : "")}
      ${cardResumoExecutivo("Pendentes", r.pendentes ?? 0, "instalação/configuração", (r.pendentes || 0) > 0 ? "warning" : "positive")}
      ${cardResumoExecutivo("Checklist", (r.progresso_checklist ?? 0) + "%", "concluído", (r.progresso_checklist || 0) >= 80 ? "positive" : "warning")}
    </div>

    <div id="implantacaoFormHostV8210">${renderFormularioImplantacaoV8210({ versao: VERSAO_ESPERADA_BADGE })}</div>

    <section class="exec-panel implant-list-panel">
      <div class="exec-panel-title">
        <div>
          <h4>Notebooks da implantação</h4>
          <p>Visão por máquina, responsável, versão, status e checklist de homologação.</p>
        </div>
        <span class="exec-chip">${notebooks.length} registro(s)</span>
      </div>
      <div class="implant-list">
        ${notebooks.length ? notebooks.map(renderNotebookImplantacaoV8210).join("") : `<div class="exec-empty">Nenhum notebook cadastrado ainda. Use o formulário acima para criar o primeiro registro.</div>`}
      </div>
    </section>
  `;
}

async function carregarPainelImplantacaoV8210() {
  const paginaEsperada = "painel_implantacao";
  const tokenEsperado = tokenAtualPagina();
  if (qs("statusCarga")) qs("statusCarga").textContent = "Carregando Painel de Implantação...";
  const { controller, timer } = criarControllerNavegacao(12000);
  try {
    const resp = await fetch("/api/painel-implantacao", { signal: controller.signal });
    const json = await resp.json();
    if (!requisicaoAindaValida(paginaEsperada, tokenEsperado)) return;
    if (!resp.ok || json.ok === false) throw new Error(json.erro || `Erro HTTP ${resp.status}`);
    renderPainelImplantacaoV8210(json);
    if (qs("statusCarga")) qs("statusCarga").textContent = `${json.resumo?.total ?? 0} notebook(s) · ${json.resumo?.homologados ?? 0} homologado(s) · ${json.resumo?.progresso_checklist ?? 0}% checklist.`;
  } catch (e) {
    if (!requisicaoAindaValida(paginaEsperada, tokenEsperado)) return;
    const painel = qs("dashboardExecutivoPremium");
    if (painel) painel.innerHTML = `<div class="exec-empty">Erro ao carregar Painel de Implantação: ${valorSeguro(e.message, "falha desconhecida")}</div>`;
    if (qs("statusCarga")) qs("statusCarga").textContent = "Erro ao carregar Painel de Implantação: " + e.message;
  } finally {
    finalizarControllerNavegacao(controller, timer);
  }
}


// ============================================================
// v8.21.38 - Diagnóstico Pós-Instalação
// ============================================================
function classeDiagnosticoInstalacaoV82128(nivel) {
  nivel = String(nivel || "ok");
  if (nivel === "erro") return "danger";
  if (nivel === "atencao") return "warning";
  return "ok";
}

function renderItemDiagnosticoInstalacaoV82128(item) {
  const cls = classeDiagnosticoInstalacaoV82128(item.nivel);
  const acao = item.acao ? `<small>${valorSeguro(item.acao, "")}</small>` : "";
  return `
    <div class="install-diag-item ${cls}">
      <div>
        <strong>${valorSeguro(item.nome, "Item")}</strong>
        <span>${valorSeguro(item.detalhe, "-")}</span>
        ${acao}
      </div>
      <b>${valorSeguro(item.estado, item.ok ? "OK" : "Atenção")}</b>
    </div>
  `;
}

function renderDiagnosticoInstalacaoV82128(payload) {
  const painel = qs("dashboardExecutivoPremium");
  if (!painel) return;
  const resumo = payload.resumo || {};
  const nivel = payload.nivel || "ok";
  const itens = payload.itens || [];
  const passos = payload.proximos_passos || [];
  painel.innerHTML = `
    <div class="install-diag-hero ${classeDiagnosticoInstalacaoV82128(nivel)}">
      <div>
        <h3>${payload.pronto ? "Pronto para Operar" : "Atenção na instalação"}</h3>
      </div>
      <button class="btn" onclick="carregarDiagnosticoInstalacaoV82128()">Atualizar</button>
    </div>

    <div class="exec-kpi-grid">
      ${cardResumoExecutivo("Itens OK", resumo.ok ?? 0, `${resumo.total ?? 0} verificados`, payload.pronto ? "positive" : "")}
      ${cardResumoExecutivo("Atenções", resumo.atencoes ?? 0, "podem operar com ressalva", (resumo.atencoes || 0) > 0 ? "warning" : "positive")}
      ${cardResumoExecutivo("Erros", resumo.erros ?? 0, "bloqueios técnicos", (resumo.erros || 0) > 0 ? "negative" : "positive")}
      ${cardResumoExecutivo("Versão", valorSeguro(payload.versao, "-"), "ativa no servidor", "positive")}
    </div>

    <section class="exec-panel install-diag-panel">
      <div class="exec-panel-title">
        <div>
          <h4>Checklist Técnico</h4>
          <p>Versão, pacote, scripts, pastas, dependências, Outlook e usuário.</p>
        </div>
        <span class="exec-chip">${valorSeguro(payload.base, "-")}</span>
      </div>
      <div class="install-diag-grid">
        ${itens.length ? itens.map(renderItemDiagnosticoInstalacaoV82128).join("") : `<div class="exec-empty">Nenhum item retornado.</div>`}
      </div>
    </section>

    <section class="exec-panel install-diag-panel">
      <div class="exec-panel-title">
        <div>
          <h4>Próximos Passos</h4>
          <p>Ações sugeridas para deixar o notebook homologado.</p>
        </div>
      </div>
      <div class="install-next-list">
        ${passos.length ? passos.map(p => `<div>${valorSeguro(p, "")}</div>`).join("") : `<div class="exec-alert ok"><strong>Sem pendências críticas</strong><span>Instalação pronta para a rotina operacional.</span></div>`}
      </div>
    </section>
  `;
}

function abrirDiagnosticoInstalacaoV82128(el) {
  iniciarTransicaoPagina("diagnostico_instalacao");
  paginaAtual = "diagnostico_instalacao";
  marcarPaginaAtualClasse();
  setVisibilidadeDashboardExecutivoPremium(true);
  dadosConsolidadoAtual = [];

  document.querySelectorAll(".menu button").forEach(b => b.classList.remove("active"));
  if (el) el.classList.add("active");

  if (qs("tituloPagina")) qs("tituloPagina").textContent = "Diagnóstico";
  if (qs("subtituloPagina")) qs("subtituloPagina").textContent = "Validação pós-instalação do notebook.";
  if (qs("tituloTabela")) qs("tituloTabela").textContent = "Diagnóstico Pós-Instalação";
  if (qs("statusCarga")) qs("statusCarga").textContent = "Verificando instalação...";

  if (qs("filtros")) qs("filtros").classList.add("hidden");
  if (qs("filtrosDashboard")) qs("filtrosDashboard").classList.add("hidden");
  if (qs("boxFiltroTipo")) qs("boxFiltroTipo").classList.add("hidden");
  if (qs("controleOrdenacaoBancoHoras")) qs("controleOrdenacaoBancoHoras").classList.add("hidden");
  if (qs("tabelaDashboard")) qs("tabelaDashboard").classList.add("hidden");
  if (qs("tabelaConsolidado")) qs("tabelaConsolidado").classList.add("hidden");
  ocultarFiltroDataPrincipalPremium();

  carregarDiagnosticoInstalacaoV82128();
}

async function carregarDiagnosticoInstalacaoV82128() {
  const paginaEsperada = "diagnostico_instalacao";
  const tokenEsperado = tokenAtualPagina();
  const painel = qs("dashboardExecutivoPremium");
  if (painel) painel.innerHTML = `<div class="admin-panel"><h3>Verificando instalação...</h3></div>`;
  try {
    const resp = await fetch("/api/diagnostico-pos-instalacao");
    const json = await resp.json();
    if (!resp.ok || json.ok === false) throw new Error(json.erro || `Erro HTTP ${resp.status}`);
    if (!requisicaoAindaValida(paginaEsperada, tokenEsperado)) return;
    renderDiagnosticoInstalacaoV82128(json);
    if (qs("statusCarga")) qs("statusCarga").textContent = json.pronto ? "Instalação pronta para operar." : "Instalação com pontos de atenção.";
  } catch (e) {
    if (painel) painel.innerHTML = `<div class="exec-empty">Erro ao carregar diagnóstico: ${valorSeguro(e.message, "falha desconhecida")}</div>`;
    if (qs("statusCarga")) qs("statusCarga").textContent = "Erro ao carregar diagnóstico: " + e.message;
  }
}


// ============================================================
// v8.21.38 - Central de Operação / Log Amigável
// ============================================================
function abrirCentralOperacaoV8190(el) {
  iniciarTransicaoPagina("central_operacao");
  paginaAtual = "central_operacao";
  marcarPaginaAtualClasse();
  setVisibilidadeDashboardExecutivoPremium(true);
  dadosConsolidadoAtual = [];

  document.querySelectorAll(".menu button").forEach(b => b.classList.remove("active"));
  if (el) el.classList.add("active");

  if (qs("tituloPagina")) qs("tituloPagina").textContent = "Central de Operação";
  if (qs("subtituloPagina")) qs("subtituloPagina").textContent = "Log amigável, saúde, alertas inteligentes e últimos eventos do robô.";
  if (qs("tituloTabela")) qs("tituloTabela").textContent = "Central de Operação";
  if (qs("statusCarga")) qs("statusCarga").textContent = "Carregando Central de Operação...";

  if (qs("filtros")) qs("filtros").classList.add("hidden");
  if (qs("filtrosDashboard")) qs("filtrosDashboard").classList.add("hidden");
  if (qs("boxFiltroTipo")) qs("boxFiltroTipo").classList.add("hidden");
  if (qs("controleOrdenacaoBancoHoras")) qs("controleOrdenacaoBancoHoras").classList.add("hidden");
  if (qs("tabelaDashboard")) qs("tabelaDashboard").classList.add("hidden");
  if (qs("tabelaConsolidado")) qs("tabelaConsolidado").classList.add("hidden");
  ocultarFiltroDataPrincipalPremium();

  carregarCentralOperacaoV8190();
}

function renderEventoOperacaoV8190(e) {
  const status = String(e.status || "INFO").toUpperCase();
  const cls = status === "OK" ? "ok" : (status === "ERRO" || status === "FALHA" ? "danger" : "info");
  const segundos = e.segundos !== null && e.segundos !== undefined ? ` · ${e.segundos}s` : "";
  const resumo = e.resumo ? `<small>${valorSeguro(e.resumo, "")}</small>` : "";
  return `
    <div class="op-event ${cls}">
      <div>
        <strong>${valorSeguro(e.tipo, "Operação")}</strong>
        <span>${valorSeguro(e.mensagem, "Sem mensagem registrada.")}</span>
        ${resumo}
      </div>
      <div class="op-event-meta">
        <b>${status}</b>
        <span>${valorSeguro(e.quando, "-")}${segundos}</span>
        <span>Entrada: ${valorSeguro(e.pdfs_entrada, 0)} PDF(s)</span>
      </div>
    </div>
  `;
}


function nivelClasseSaudeV8191(nivel) {
  nivel = String(nivel || "ok");
  if (nivel === "erro") return "danger";
  if (nivel === "atencao") return "warning";
  if (nivel === "sem_historico") return "info";
  return "ok";
}

function renderDiagnosticoSaudeV8191(d) {
  const cls = nivelClasseSaudeV8191(d.nivel || "ok");
  return `
    <div class="op-health-item ${cls}">
      <div>
        <strong>${valorSeguro(d.nome, "Diagnóstico")}</strong>
        <span>${valorSeguro(d.detalhe, "-")}</span>
      </div>
      <b>${valorSeguro(d.estado, "OK")}</b>
    </div>
  `;
}

function renderMedidorSaudeV8191(saude) {
  const score = Number(saude.score ?? 0);
  const nivel = saude.nivel || "ok";
  return `
    <div class="op-health-meter ${nivel}">
      <div class="op-health-score">
        <strong>${Math.max(0, Math.min(100, score))}</strong>
        <span>saúde</span>
      </div>
      <div class="op-health-meter-copy">
        <h4>${valorSeguro(saude.titulo, "Saúde do Robô")}</h4>
        <p>${valorSeguro(saude.descricao, "Diagnóstico operacional disponível.")}</p>
        <small>Outlook: ${valorSeguro(saude.politica_outlook, "-")} · Entrada: ${valorSeguro(saude.pdfs_entrada, 0)} PDF(s) · Erro: ${valorSeguro(saude.pdfs_erro, 0)} PDF(s)</small>
      </div>
    </div>
  `;
}


function classePrioridadeAlertaV8192(prioridade) {
  prioridade = String(prioridade || "info").toLowerCase();
  if (prioridade === "alto") return "danger";
  if (prioridade === "medio") return "warning";
  if (prioridade === "baixo") return "ok";
  return "info";
}

function renderAlertaInteligenteV8192(a) {
  const cls = classePrioridadeAlertaV8192(a.prioridade);
  const prioridade = String(a.prioridade || "info").toUpperCase();
  return `
    <div class="op-smart-alert ${cls}">
      <div class="op-smart-alert-main">
        <div class="op-smart-alert-head">
          <strong>${valorSeguro(a.titulo, "Alerta operacional")}</strong>
          <span>${valorSeguro(a.categoria, "Operação")}</span>
        </div>
        <p>${valorSeguro(a.descricao, "Há um sinal operacional para acompanhamento.")}</p>
        <small><b>Ação sugerida:</b> ${valorSeguro(a.acao, "Revisar o item indicado.")}</small>
      </div>
      <b class="op-smart-priority ${cls}">${prioridade}</b>
    </div>
  `;
}


function renderLinhaLogAmigavelV8193(l) {
  const cls = l.classe || "info";
  return `
    <div class="op-friendly-line ${cls}">
      <div class="op-friendly-dot"></div>
      <div class="op-friendly-copy">
        <div class="op-friendly-head">
          <strong>${valorSeguro(l.titulo, "Evento registrado")}</strong>
          <span>${valorSeguro(l.quando_amigavel, "-")}</span>
        </div>
        <p>${valorSeguro(l.frase, "Registro operacional disponível.")}</p>
        ${l.detalhe ? `<small>${valorSeguro(l.detalhe, "")}</small>` : ""}
      </div>
      <b>${valorSeguro(l.status, "INFO")}</b>
    </div>
  `;
}

function renderLogAmigavelV8193(log) {
  log = log || {};
  const linhas = log.linhas || [];
  const metricas = log.metricas || {};
  const passos = log.proximos_passos || [];
  const tom = log.tom || "info";
  return `
    <section class="exec-panel op-friendly-panel ${tom}">
      <div class="exec-panel-title">
        <div>
          <h4>Log Amigável</h4>
          <p>Resumo em linguagem simples para operação diária.</p>
        </div>
        <span class="exec-chip">${valorSeguro(metricas.eventos, 0)} evento(s)</span>
      </div>
      <div class="op-friendly-summary ${tom}">
        <div>
          <strong>${valorSeguro(log.titulo, "Resumo operacional")}</strong>
          <p>${valorSeguro(log.resumo, "Sem resumo disponível.")}</p>
        </div>
        <div class="op-friendly-metrics">
          <span><b>${valorSeguro(metricas.sucessos, 0)}</b> sucesso(s)</span>
          <span><b>${valorSeguro(metricas.erros, 0)}</b> erro(s)</span>
        </div>
      </div>
      <div class="op-friendly-next">
        <strong>Próximos passos</strong>
        ${passos.length ? `<ul>${passos.map(p => `<li>${valorSeguro(p, "Acompanhar operação.")}</li>`).join("")}</ul>` : `<p>Manter acompanhamento normal.</p>`}
      </div>
      <div class="op-friendly-list">
        ${linhas.length ? linhas.map(renderLinhaLogAmigavelV8193).join("") : `<div class="exec-empty">Nenhum evento amigável disponível ainda.</div>`}
      </div>
    </section>
  `;
}

function renderCentralOperacaoV8190(payload) {
  const painel = qs("dashboardExecutivoPremium");
  if (!painel) return;
  const k = payload.kpis || {};
  const saude = payload.saude || {};
  const eventos = payload.eventos || [];
  const job = payload.processamento_integrado || {};
  const alertas = saude.alertas || [];
  const alertasInteligentes = payload.alertas_inteligentes || [];
  const resumoAlertas = payload.resumo_alertas || {};
  const logAmigavel = payload.log_amigavel || {};
  const nivel = saude.nivel || "ok";

  painel.innerHTML = `
    <div class="op-hero ${nivel}">
      <div>
        <h3>Central de Operação</h3>
      </div>
      <button class="btn" onclick="carregarCentralOperacaoV8190()">Atualizar</button>
    </div>

    <div class="exec-kpi-grid">
      ${cardResumoExecutivo("Alertas", resumoAlertas.total ?? 0, "inteligentes", (resumoAlertas.altos || 0) > 0 ? "negative" : "")}
      ${cardResumoExecutivo("Críticos", resumoAlertas.altos ?? 0, "prioridade alta", (resumoAlertas.altos || 0) > 0 ? "negative" : "positive")}
      ${cardResumoExecutivo("Erros", k.erros ?? 0, "histórico recente", (k.erros || 0) > 0 ? "negative" : "")}
      ${cardResumoExecutivo("PDFs entrada", k.pdfs_entrada ?? 0, "aguardando processamento", (k.pdfs_entrada || 0) > 0 ? "warning" : "positive")}
    </div>

    <section class="exec-panel op-smart-alert-panel">
      <div class="exec-panel-title">
        <div>
          <h4>Alertas Operacionais Inteligentes</h4>
          <p>Prioridades, motivo do alerta e ação sugerida para a próxima decisão.</p>
        </div>
        <span class="exec-chip">${resumoAlertas.altos ?? 0} crítico(s)</span>
      </div>
      <div class="op-smart-alert-list">
        ${alertasInteligentes.length ? alertasInteligentes.map(renderAlertaInteligenteV8192).join("") : `<div class="exec-empty">Nenhum alerta inteligente disponível.</div>`}
      </div>
    </section>

    ${renderLogAmigavelV8193(logAmigavel)}

    <section class="exec-panel op-health-panel">
      <div class="exec-panel-title">
        <div>
          <h4>Saúde do Robô</h4>
          <p>Diagnóstico ampliado com base em pastas, arquivos-chave, configuração e histórico recente.</p>
        </div>
        <span class="exec-chip">${valorSeguro(payload.gerado_em, "-")}</span>
      </div>
      ${renderMedidorSaudeV8191(saude)}
      <div class="op-health-grid">
        ${(saude.diagnosticos || []).length ? (saude.diagnosticos || []).map(renderDiagnosticoSaudeV8191).join("") : `<div class="exec-empty">Nenhum diagnóstico disponível.</div>`}
      </div>
      <div class="exec-alert-list op-alerts-compact">
        ${alertas.length ? alertas.map(a => `
          <div class="exec-alert ${a.nivel === "alto" ? "danger" : a.nivel === "medio" ? "warning" : "ok"}">
            <strong>${valorSeguro(a.titulo, "Alerta")}</strong>
            <span>${valorSeguro(a.descricao, "")}</span>
          </div>
        `).join("") : `<div class="exec-alert ok"><strong>Sem alertas ativos</strong><span>Não há sinal operacional relevante neste momento.</span></div>`}
      </div>
    </section>

    <section class="exec-panel">
      <div class="exec-panel-title">
        <div>
          <h4>Processamento Integrado</h4>
          <p>Último estado conhecido do job em segundo plano.</p>
        </div>
        <span class="exec-chip">${valorSeguro(job.status, "ocioso")}</span>
      </div>
      <div class="op-job-box">
        <strong>${valorSeguro(job.fase, "Aguardando")}</strong>
        <span>${valorSeguro(job.mensagem, "Nenhum processamento em andamento.")}</span>
        <small>Início: ${valorSeguro(job.iniciado_em, "-")} · Fim: ${valorSeguro(job.finalizado_em, "-")} · ${valorSeguro(job.segundos, 0)}s</small>
      </div>
    </section>

    <section class="exec-panel">
      <div class="exec-panel-title">
        <div>
          <h4>Últimas Operações</h4>
          <p>Diário de bordo operacional gerado a partir da v8.19.0.</p>
        </div>
      </div>
      <div class="op-event-list">
        ${eventos.length ? eventos.map(renderEventoOperacaoV8190).join("") : `<div class="exec-empty">Nenhum evento registrado ainda. Execute uma importação ou processamento para iniciar o histórico.</div>`}
      </div>
    </section>
  `;
}

async function carregarCentralOperacaoV8190() {
  const paginaEsperada = "central_operacao";
  const tokenEsperado = tokenAtualPagina();
  if (qs("statusCarga")) qs("statusCarga").textContent = "Carregando Central de Operação...";
  const { controller, timer } = criarControllerNavegacao(12000);
  try {
    const resp = await fetch("/api/central-operacao?limite=30", { signal: controller.signal });
    const json = await resp.json();
    if (!requisicaoAindaValida(paginaEsperada, tokenEsperado)) return;
    if (!resp.ok || json.ok === false) {
      renderCentralOperacaoV8190({});
      if (qs("statusCarga")) qs("statusCarga").textContent = json.erro || `Erro HTTP ${resp.status}`;
      return;
    }
    renderCentralOperacaoV8190(json);
    if (qs("statusCarga")) qs("statusCarga").textContent = `${json.kpis?.eventos ?? 0} evento(s) · ${json.resumo_alertas?.total ?? 0} alerta(s) · Log Amigável atualizado.`;
  } catch (e) {
    if (!requisicaoAindaValida(paginaEsperada, tokenEsperado)) return;
    renderCentralOperacaoV8190({});
    if (qs("statusCarga")) qs("statusCarga").textContent = "Erro ao carregar Central de Operação: " + e.message;
  } finally {
    finalizarControllerNavegacao(controller, timer);
  }
}

// v8.21.9 - Verificacao da versao realmente ativa no navegador/backend.
const VERSAO_ESPERADA_BADGE = "v8.21.38";
let versaoAtivaV8215 = null;

function formatarVersaoAtivaV8215(info) {
  if (!info || !info.ok) return "Versão ativa indisponível";
  return [
    `${info.titulo || info.versao || "Tempo Fechado"}`,
    `Iniciado em: ${info.iniciado_em || "-"}`,
    `Host: ${info.host || "-"}`,
    `Máquina: ${info.maquina || "-"}`,
    `Usuário Windows: ${info.usuario_windows || "-"}`,
    `Pasta do app: ${info.pasta_app || "-"}`,
    `Backend: ${info.arquivo_backend || "-"}`,
    `Executável: ${info.executavel || "-"}`,
    `Modo EXE: ${info.python_frozen ? "Sim" : "Não"}`
  ].join("\n");
}

async function carregarVersaoAtivaV8215() {
  const badge = qs("versaoAtivaBadge");
  try {
    if (badge) badge.textContent = "Verificando versão...";
    const resp = await fetch("/api/versao", { cache: "no-store" });
    const json = await resp.json();
    versaoAtivaV8215 = json;
    if (badge) {
      badge.textContent = json.versao ? `${json.versao} ativa` : "Versão ativa";
      badge.title = formatarVersaoAtivaV8215(json);
      badge.classList.toggle("version-active-warning", json.versao !== VERSAO_ESPERADA_BADGE);
    }
  } catch (e) {
    versaoAtivaV8215 = { ok: false, erro: e.message };
    if (badge) {
      badge.textContent = "Versão não identificada";
      badge.title = "Não foi possível consultar /api/versao: " + e.message;
      badge.classList.add("version-active-warning");
    }
  }
}

function mostrarVersaoAtivaV8215() {
  if (versaoAtivaV8215 && versaoAtivaV8215.ok) {
    alert(formatarVersaoAtivaV8215(versaoAtivaV8215));
  } else {
    alert("Não foi possível identificar a versão ativa. Tente recarregar a página ou verifique se há uma instância antiga do Tempo Fechado aberta.");
  }
}

document.addEventListener("DOMContentLoaded", carregarVersaoAtivaV8215);
