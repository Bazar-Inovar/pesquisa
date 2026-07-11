const CONFIG = {
  fileUrl: "/produtos.xlsx"
};

const pesquisa = document.getElementById("pesquisa");
const resultado = document.getElementById("resultado");
const loading = document.getElementById("loading");
const app = document.getElementById("app");
const statusTexto = document.getElementById("statusTexto");
const semResultado = document.getElementById("semResultado");

let produtos = [];

function atualizarStatus(t){
  if(statusTexto) statusTexto.textContent = t;
}

function parseNumber(v){
  if(v == null || v === "") return 0;
  let s = String(v).replace("R$", "").trim();
  const hasComma = s.includes(",");
  const hasDot = s.includes(".");
  if(hasComma && hasDot) s = s.replace(/\./g, "").replace(",", ".");
  else if(hasComma) s = s.replace(",", ".");
  return Number(s) || 0;
}

function normalizeKey(s){
  return String(s||"").toLowerCase().normalize("NFD")
   .replace(/[\u0300-\u036f]/g,"")
   .replace(/\s+/g,"")
   .replace(/[^a-z0-9_]/g,"");
}

function normalizeRows(rows){
  return rows.map(r=>{
    const o={};
    Object.keys(r).forEach(k=>o[normalizeKey(k)] = r[k]);
    o.valor = parseNumber(o.valor);
    o.quantidade = parseInt(o.quantidade||0,10);
    if(!o.produto && o.nome) o.produto = o.nome;
    if(!o.codigo && o.codigoproduto) o.codigo = o.codigoproduto;
    return o;
  });
}

async function carregarProdutos(){
  atualizarStatus("Baixando planilha...");
  const response = await fetch(CONFIG.fileUrl);
  if(!response.ok) throw new Error("Erro ao baixar arquivo");
  const buffer = await response.arrayBuffer();
  const wb = XLSX.read(buffer, {type: "array"});
  const ws = wb.Sheets[wb.SheetNames[0]];
  produtos = normalizeRows(XLSX.utils.sheet_to_json(ws));
  atualizarStatus(`${produtos.length} produtos carregados.`);
}

function renderizarProdutos(lista){
  if(!lista.length){
    resultado.innerHTML = "";
    if(semResultado) semResultado.classList.remove("hidden");
    return;
  }
  if(semResultado) semResultado.classList.add("hidden");

  resultado.innerHTML = lista.map(p=>`
  <div class="produto">
    <div class="produto-topo">
      <h3>${p.produto || "Sem nome"}</h3>
      <strong>R$ ${Number(p.valor||0).toFixed(2).replace(".",",")}</strong>
    </div>
    <div class="detalhes">
      <p>Qtd: ${p.quantidade}</p> |
      <p>Cód: ${p.codigo || "-"}</p>
    </div>
  </div>`).join("");
}

pesquisa?.addEventListener("input",()=>{
 const t = pesquisa.value.toLowerCase().trim();
 if(!t){resultado.innerHTML=""; if(semResultado) semResultado.classList.add("hidden"); return;}
 const filtrados = produtos.filter(p=>
   (p.produto||"").toLowerCase().includes(t) ||
   String(p.codigo||"").toLowerCase().includes(t)
 ).slice(0,30);
 renderizarProdutos(filtrados);
});

async function init(){
 loading?.classList.remove("hidden");
 app?.classList.add("hidden");
 try{
   await carregarProdutos();
 }catch(e){
   atualizarStatus("Erro: " + e.message);
   console.error(e);
 }finally{
   loading?.classList.add("hidden");
   app?.classList.remove("hidden");
 }
}

init();
