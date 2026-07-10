// ---------- CONFIGURAÇÃO ----------
const CONFIG = {
  useApi: true,                    // true para buscar remotamente; false para usar cache apenas
  apiType: "google_drive",         // Tipo de API definido para o Google Drive
  
  // Google Drive API Configurações
  apiKey: "AIzaSyA405vJpn5K60-mZhWTUAu5Y1Kmt8rVJi8", // Sua chave de API com permissão para Drive API
  folderId: "1-PE0gn6fxi82vp3Fd01gzfkW3UigCc8o",     // ID da pasta do Drive inserido com sucesso
  
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
let gapiCarregado = false;

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
    Object.keys(row).forEach(k => {
      const nk = normalizeKey(k);
      obj[nk] = row[k];
    });
    if (obj.valor) obj.valor = parseNumber(obj.valor);
    if (obj.quantidade) obj.quantidade = parseInt(String(obj.quantidade).replace(",", ""), 10) || 0;
    if (obj.codigo === undefined && obj.codigoproduto !== undefined) obj.codigo = obj.codigoproduto;
    if (obj.produto === undefined && obj.nome !== undefined) obj.produto = obj.nome;
    return obj;
  });
}

// ---------- INJETOR AUTOMÁTICO DE SCRIPT ----------
function carregarScriptExterno(url) {
  return new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = url;
    script.async = true;
    script.onload = () => resolve();
    script.onerror = () => reject(new Error(`Falha ao injetar script: ${url}`));
    document.head.appendChild(script);
  });
}

// ---------- INICIALIZADOR DO GOOGLE API (GAPI) ----------
async function inicializarGapi() {
  if (gapiCarregado) return;

  atualizarStatus("Injetando dependências do Google e Excel...");
  
  // Injeta a biblioteca XLSX se ela não existir no escopo global
  if (typeof XLSX === "undefined") {
    await carregarScriptExterno("https://cloudflare.com");
  }

  // Injeta a biblioteca GAPI oficial correta se ela não existir
  if (typeof gapi === "undefined") {
    await carregarScriptExterno("https://google.com");
  }

  return new Promise((resolve, reject) => {
    gapi.load("client", () => {
      gapi.client.init({
        apiKey: CONFIG.apiKey,
        discoveryDocs: ["https://googleapis.com"],
      })
      .then(() => {
        gapiCarregado = true;
        resolve();
      })
      .catch(err => reject(err));
    });
  });
}

// ---------- FETCH: Google Drive (Busca, Download e Conversão XLSX) ----------
async function fetchFromGoogleDrive() {
  if (!CONFIG.apiKey || !CONFIG.folderId) {
    throw new Error("apiKey ou folderId não configurados para o Google Drive");
  }

  // 1) Garante injeção e inicialização do ecossistema Google
  await inicializarGapi();

  atualizarStatus("Localizando PRODUTOS.xlsx no Drive...");
  
  // 2) Procura o arquivo pelo nome exato dentro da pasta configurada
  const q = `name = 'PRODUTOS.xlsx' and '${CONFIG.folderId}' in parents and trashed = false`;
  const listaResponse = await gapi.client.drive.files.list({
    q: q,
    fields: "files(id, name)",
    pageSize: 1
  });

  const arquivos = listaResponse.result.files;
  if (!arquivos || arquivos.length === 0) {
    throw new Error("Arquivo 'PRODUTOS.xlsx' não foi encontrado na pasta informada.");
  }

  // Ajustado com [0] para pegar o primeiro arquivo retornado
  const fileId = arquivos[0].id;
  atualizarStatus("Baixando arquivo Excel...");

  // 3) Faz o download do arquivo binário usando a API Key pública
  const downloadUrl = `https://googleapis.com{fileId}?alt=media&key=${CONFIG.apiKey}`;
  const response = await fetch(downloadUrl);
  
  if (!response.ok) {
    throw new Error("Falha ao baixar o arquivo binário do Drive.");
  }

  const arrayBuffer = await response.arrayBuffer();
  
  atualizarStatus("Convertendo Excel para JSON...");

  // 4) Faz a leitura da estrutura do Excel puro diretamente no Navegador
  const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: "array" });
  
  // Ajustado com [0] para obter o nome da primeira aba disponível
  const primeiraAbaNome = workbook.SheetNames[0]; 
  const aba = workbook.Sheets[primeiraAbaNome];
  
  // Transforma as linhas estruturadas em formato JSON legível
  const jsonBruto = XLSX.utils.sheet_to_json(aba);

  return normalizeRows(jsonBruto);
}

// ---------- FUNÇÃO PRINCIPAL ----------
async function carregarProdutos(atualizar = false) {
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

  if (!CONFIG.useApi) {
    atualizarStatus("Busca remota desativada. Ative CONFIG.useApi para buscar dados.");
    produtos = [];
    return produtos;
  }

  atualizarStatus("Buscando produtos no servidor...");
  let dados = [];
  try {
    if (CONFIG.apiType === "google_drive") {
      dados = await fetchFromGoogleDrive();
    } else {
      throw new Error("CONFIG.apiType inválido. Use 'google_drive'.");
    }
  } catch (err) {
    atualizarStatus(`Erro ao buscar dados: ${err.message}`);
    throw err;
  }

  if (!dados || dados.length === 0) {
    atualizarStatus("Nenhum dado retornado do arquivo.");
    produtos = [];
    localStorage.setItem("produtos", JSON.stringify({ produtos, data: Date.now() }));
    return produtos;
  }

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
}

function esconderLoading() {
  if (loading) loading.classList.add("hidden");
}
