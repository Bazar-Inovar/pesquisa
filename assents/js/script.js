// ---------- CONFIGURAÇÃO (preencha amanhã) ----------
const CONFIG = {
  useApi: true,                    // true para buscar remotamente; false para usar cache apenas
  apiType: "json_endpoint",        // "google_sheets" or "json_endpoint"
  // Google Sheets API v4 (se apiType === "google_sheets")
  sheetId: "1-PE0gn6fx82vp3Fd01gzfkW3UigCc80",
  apiKey: "AIzaSyA405vJpn5K60-mZhWTUAu5Y1Kmt8rVJi8",
  range: "A1:Z",                   // range a ser lido na Sheets API
  // Endpoint JSON (se apiType === "json_endpoint")
  endpointUrl: "https://api.seuservidor.com/produtos", // deve retornar array de objetos
  // Cache
  cacheTime: 1000 * 60 * 5         // 5 minutos
};

// ---------- ELEMENTOS DOM ----------
const pesquisa = document.getElementById("pesquisa");
const resultado = document.getElementById("resultado");
const loading = document.getElementById("loading");
const app = document.getElementById("app");
const statusTexto = document.getElementById("statusTexto");
const btnAtualizar = document.getElementById("btnAtualizar");
const semResultado = document.getElementById("semResultado");

// ---------- ESTADO GLOBAL ----------
let produtos = [];
let ultimaAtualizacao = null;

// ---------- HELPERS ----------
function normalizeKey(str) {
  return String(str || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, "")
    .replace(/[^a-z0-9_]/g, "");
}

function parseNumber(value) {
  if (value === null || value === undefined || value === "") return 0;
  const s = String(value).replace(/\s/g, "").replace("R$", "").replace(/\./g, "").replace(",", ".");
  const n = parseFloat(s);
  return Number.isNaN(n) ? 0 : n;
}

function atualizarStatus(texto) {
  if (statusTexto) statusTexto.textContent = texto;
  console.log(texto);
}

// Normaliza linhas vindas da API para as chaves esperadas
function normalizeRows(rows) {
  return (rows || []).map(row => {
    const obj = {};
    // Se row for array de valores com header separado, não coberto aqui.
    // Assume row é objeto { "Código": "...", "Produto": "...", ... } ou já normalizado.
    Object.keys(row).forEach(k => {
      const nk = normalizeKey(k);
      obj[nk] = row[k];
    });
    // normalizações comuns
    if (obj.valor) obj.valor = parseNumber(obj.valor);
    if (obj.quantidade) obj.quantidade = parseInt(String(obj.quantidade).replace(",", ""), 10) || 0;
    // mapeamentos alternativos
    if (obj.codigo === undefined && obj.codigoproduto !== undefined) obj.codigo = obj.codigoproduto;
    if (obj.produto === undefined && obj.nome !== undefined) obj.produto = obj.nome;
    return obj;
  });
}

// ---------- FETCH: Google Sheets API v4 ----------
async function fetchFromGoogleSheets() {
  if (!CONFIG.sheetId || !CONFIG.apiKey) {
    throw new Error("sheetId ou apiKey não configurados para Google Sheets");
  }
  const url = `https://sheets.googleapis.com/v4/spreadsheets/${CONFIG.sheetId}/values/${encodeURIComponent(CONFIG.range)}?key=${CONFIG.apiKey}`;
  const res = await fetch(url);
  if (!res.ok) throw new Error("Erro ao acessar Google Sheets API");
  const json = await res.json();
  const [headerRow, ...rows] = json.values || [];
  if (!headerRow) return [];
  const headers = headerRow.map(h => normalizeKey(h));
  const data = rows.map(r => {
    const obj = {};
    headers.forEach((h, i) => {
      obj[h] = r[i] ?? "";
    });
    return obj;
  });
  return normalizeRows(data);
}

// ---------- FETCH: Endpoint JSON genérico ----------
async function fetchFromJsonEndpoint() {
  if (!CONFIG.endpointUrl) throw new Error("endpointUrl não configurado");
  const res = await fetch(CONFIG.endpointUrl, { cache: "no-store" });
  if (!res.ok) throw new Error("Erro ao acessar endpoint JSON");
  const json = await res.json();
  // espera-se que json seja um array de objetos
  if (!Array.isArray(json)) {
    throw new Error("Resposta do endpoint não é um array");
  }
  return normalizeRows(json);
}

// ---------- FUNÇÃO PRINCIPAL (mantendo nome) ----------
/*
  carregarProdutos(atualizar = false)
  - busca dados remotos quando CONFIG.useApi = true
  - usa cache local quando disponível e dentro do tempo
*/
async function carregarProdutos(atualizar = false) {
  // 1) Verifica cache
  if (!atualizar) {
    const cache = localStorage.getItem("produtos");
    if (cache) {
      try {
        const dadosCache = JSON.parse(cache);
        if (Date.now() - dadosCache.data < CONFIG.cacheTime) {
          produtos = dadosCache.produtos || [];
          ultimaAtualizacao = new Date(dadosCache.data);
          atualizarStatus("Produtos carregados do cache.");
          return produtos;
        }
      } catch (e) {
        console.warn("Cache inválido, ignorando.");
      }
    }
  }

  // 2) Se não usar API, apenas limpa estado e retorna (sem upload local)
  if (!CONFIG.useApi) {
    atualizarStatus("Busca remota desativada. Ative CONFIG.useApi para buscar dados.");
    produtos = [];
    return produtos;
  }

  // 3) Buscar via API conforme tipo configurado
  atualizarStatus("Buscando produtos no servidor...");
  let dados = [];
  try {
    if (CONFIG.apiType === "google_sheets") {
      dados = await fetchFromGoogleSheets();
    } else if (CONFIG.apiType === "json_endpoint") {
      dados = await fetchFromJsonEndpoint();
    } else {
      throw new Error("CONFIG.apiType inválido. Use 'google_sheets' ou 'json_endpoint'.");
    }
  } catch (err) {
    atualizarStatus(`Erro ao buscar dados: ${err.message}`);
    throw err;
  }

  if (!dados || dados.length === 0) {
    atualizarStatus("Nenhum dado retornado da API.");
    produtos = [];
    localStorage.setItem("produtos", JSON.stringify({ produtos, data: Date.now() }));
    return produtos;
  }

  // 4) Salva estado e cache
  produtos = dados;
  ultimaAtualizacao = new Date();
  localStorage.setItem("produtos", JSON.stringify({ produtos, data: Date.now() }));
  atualizarStatus("Produtos carregados com sucesso.");
  return produtos;
}

// ---------- RENDERIZA PRODUTOS ----------
function renderizarProdutos(lista) {
  resultado.innerHTML = "";

  if (!lista || lista.length === 0) {
    semResultado.classList.remove("hidden");
    return;
  }

  semResultado.classList.add("hidden");

  let html = "";

  lista.forEach(produto => {
    const nome = produto.produto || "";
    const valor = Number(produto.valor || 0);
    const quantidade = produto.quantidade ?? "";
    const codigo = produto.codigo ?? "";

    html += `
      <div class="produto">
        <nav class="produto-topo"> 
          <h3 class="produto-nome">${nome}</h3>
          <strong class="produto-preco">R$ ${valor.toFixed(2)}</strong>
        </nav>
        <nav class="detalhes"> 
          <p>Qtd: ${quantidade}</p> |
          <p>Cód: ${codigo}</p> 
        </nav> 
      </div>
    `;
  });

  resultado.innerHTML = html;
}

// ---------- PESQUISA ----------
if (pesquisa) {
  pesquisa.addEventListener("input", () => {
    const texto = pesquisa.value.toLowerCase().trim();

    if (texto === "") {
      resultado.innerHTML = "";
      semResultado.classList.add("hidden");
      return;
    }

    const filtrados = produtos.filter(produto => {
      const nome = (produto.produto || "").toString().toLowerCase();
      const codigo = (produto.codigo || "").toString().toLowerCase();
      return nome.includes(texto) || codigo.includes(texto);
    });

    const limite = filtrados.slice(0, 30);
    renderizarProdutos(limite);
  });
}

// ---------- BOTÃO ATUALIZAR ----------
if (btnAtualizar) {
  btnAtualizar.addEventListener("click", () => {
    mostrarLoading();
    carregarProdutos(true)
      .then(() => esconderLoading())
      .catch(err => {
        console.error(err);
        esconderLoading();
      });
  });
}

// ---------- INICIALIZAÇÃO ----------
document.addEventListener("DOMContentLoaded", async () => {
  try {
    mostrarLoading();
    await carregarProdutos();
    esconderLoading();
    if (pesquisa) pesquisa.focus();
  } catch (err) {
    console.error(err);
    atualizarStatus("Erro ao inicializar produtos.");
    esconderLoading();
  }
});

// ---------- FUNÇÕES DE UI ----------
function mostrarLoading() {
  if (loading) loading.classList.remove("hidden");
  if (app) app.classList.add("hidden");
}
function esconderLoading() {
  if (loading) loading.classList.add("hidden");
  if (app) app.classList.remove("hidden");
}
