// SCRIPT CORRIGIDO (base)
// Ajuste apenas os valores abaixo.
const CONFIG = {
  useApi: true,
  apiKey: "SUA_API_KEY",
  folderId: "SEU_FOLDER_ID",
  cacheTime: 1000 * 60 * 5
};

const pesquisa = document.getElementById("pesquisa");
const resultado = document.getElementById("resultado");
const loading = document.getElementById("loading");
const app = document.getElementById("app");
const statusTexto = document.getElementById("statusTexto");
const btnAtualizar = document.getElementById("btnAtualizar");
const semResultado = document.getElementById("semResultado");

let produtos = [];

function atualizarStatus(t){
  if(statusTexto) statusTexto.textContent=t;
}

function parseNumber(v){
  if(v==null||v==="") return 0;
  return Number(String(v).replace("R$","").replace(/\./g,"").replace(",","."));
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
    Object.keys(r).forEach(k=>o[normalizeKey(k)]=r[k]);
    o.valor=parseNumber(o.valor);
    o.quantidade=parseInt(o.quantidade||0,10);
    if(!o.produto && o.nome) o.produto=o.nome;
    if(!o.codigo && o.codigoproduto) o.codigo=o.codigoproduto;
    return o;
  });
}

async function fetchFromGoogleDrive(){
  atualizarStatus("Localizando arquivo...");
  const q=encodeURIComponent(`name='produtos.xlsx' and '${CONFIG.folderId}' in parents and trashed=false`);
  const list=`https://www.googleapis.com/drive/v3/files?q=${q}&fields=files(id,name)&key=${CONFIG.apiKey}`;
  const r=await fetch(list);
  if(!r.ok) throw new Error("Erro na Drive API");
  const j=await r.json();
  if(!j.files.length) throw new Error("Arquivo não encontrado");
  const id=j.files[0].id;

  atualizarStatus("Baixando planilha...");
  const dl=`https://www.googleapis.com/drive/v3/files/${id}?alt=media&key=${CONFIG.apiKey}`;
  const ar=await fetch(dl);
  if(!ar.ok) throw new Error("Falha no download");

  const buf=await ar.arrayBuffer();
  const wb=XLSX.read(buf,{type:"array"});
  const ws=wb.Sheets[wb.SheetNames[0]];
  return normalizeRows(XLSX.utils.sheet_to_json(ws));
}

async function carregarProdutos(force=false){
  atualizarStatus("Carregando...");
  produtos=await fetchFromGoogleDrive();
  atualizarStatus(`${produtos.length} produtos carregados.`);
}

function renderizarProdutos(lista){
  resultado.innerHTML=lista.map(p=>`
  <div class="produto">
    <div class="produto-topo">
      <h3>${p.produto||""}</h3>
      <strong>R$ ${Number(p.valor||0).toFixed(2)}</strong>
    </div>
    <div class="detalhes">
      <p>Qtd: ${p.quantidade}</p>
      <p>Cód: ${p.codigo||""}</p>
    </div>
  </div>`).join("");
}

pesquisa?.addEventListener("input",()=>{
 const t=pesquisa.value.toLowerCase().trim();
 if(!t){resultado.innerHTML="";return;}
 renderizarProdutos(produtos.filter(p=>(p.produto||"").toLowerCase().includes(t)||(String(p.codigo||"").toLowerCase().includes(t))).slice(0,30));
});

btnAtualizar?.addEventListener("click",()=>init(true));

async function init(force=false){
 loading?.classList.remove("hidden");
 app?.classList.add("hidden");
 try{
   await carregarProdutos(force);
 }finally{
   loading?.classList.add("hidden");
   app?.classList.remove("hidden");
 }
}
document.addEventListener("DOMContentLoaded",()=>init());
