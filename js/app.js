/* SmarTree · Planejador de Propostas
   Aplicação principal: proposta, propostas salvas, configuração de publishers, usuários,
   exportação em Excel e PowerPoint. Dados e login via Supabase (seção "camada de dados"). */
(() => {
/* ---------- utilidades ---------- */
const $ = s => document.querySelector(s);
const esc = v => String(v ?? "").replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const BRL = new Intl.NumberFormat("pt-BR",{style:"currency",currency:"BRL"});
const NUM = new Intl.NumberFormat("pt-BR",{maximumFractionDigits:0});
const pct = v => (v*100).toLocaleString("pt-BR",{maximumFractionDigits:1}) + "%";
const norm = s => String(s ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase().trim();
const slug = s => norm(s).replace(/[^a-z0-9]+/g,"-").replace(/^-|-$/g,"").slice(0,60) || "publisher";
const clone = o => JSON.parse(JSON.stringify(o));
const fmtData = iso => { if(!iso) return ""; const [y,m,d] = iso.slice(0,10).split("-"); return `${d}/${m}/${y}`; };
const lsGet = k => { try { return localStorage.getItem(k); } catch(e){ return null; } };
const lsSet = (k,v) => { try { localStorage.setItem(k,v); } catch(e){} };
const parseMoney = s => { const t = String(s||"").replace(/[^\d,.-]/g,""); if(!t) return 0; const n = t.includes(",") ? t.replace(/\./g,"").replace(",",".") : t.replace(/\.(?=\d{3}(\D|$))/g,""); return Math.max(0, parseFloat(n)||0); };
/* integrações: Supabase (js/data.js) */

/* ---------- referência: estados (população Censo IBGE 2022, milhões) ---------- */
const UFS = [["AC","Acre",0.83,"N"],["AL","Alagoas",3.13,"NE"],["AP","Amapá",0.73,"N"],["AM","Amazonas",3.94,"N"],["BA","Bahia",14.14,"NE"],["CE","Ceará",8.79,"NE"],["DF","Distrito Federal",2.82,"CO"],["ES","Espírito Santo",3.83,"SE"],["GO","Goiás",7.06,"CO"],["MA","Maranhão",6.78,"NE"],["MT","Mato Grosso",3.66,"CO"],["MS","Mato Grosso do Sul",2.76,"CO"],["MG","Minas Gerais",20.54,"SE"],["PA","Pará",8.12,"N"],["PB","Paraíba",3.97,"NE"],["PR","Paraná",11.44,"S"],["PE","Pernambuco",9.06,"NE"],["PI","Piauí",3.27,"NE"],["RJ","Rio de Janeiro",16.05,"SE"],["RN","Rio Grande do Norte",3.30,"NE"],["RS","Rio Grande do Sul",10.88,"S"],["RO","Rondônia",1.58,"N"],["RR","Roraima",0.64,"N"],["SC","Santa Catarina",7.61,"S"],["SP","São Paulo",44.41,"SE"],["SE","Sergipe",2.21,"NE"],["TO","Tocantins",1.51,"N"]];
const UF = Object.fromEntries(UFS.map(([uf,nome,pop,reg]) => [uf,{uf,nome,pop,reg}]));
const POP_TOTAL = UFS.reduce((a,u)=>a+u[2],0);
const REGIOES = {SE:"Sudeste",S:"Sul",NE:"Nordeste",CO:"Centro-Oeste",N:"Norte"};
const UF_POR_NOME = Object.fromEntries(UFS.map(u => [norm(u[1]), u[0]]));
UF_POR_NOME["federal district"] = "DF"; UF_POR_NOME["sao paulo state"] = "SP";
const CIDADES = [["belo horizonte","MG"],["curitiba","PR"],["porto alegre","RS"],["recife","PE"],["salvador","BA"],["fortaleza","CE"],["brasilia","DF"],["goiania","GO"],["manaus","AM"],["belem","PA"],["florianopolis","SC"],["vitoria","ES"],["campinas","SP"],["natal","RN"],["joao pessoa","PB"],["teresina","PI"],["maceio","AL"],["aracaju","SE"],["sao luis","MA"],["cuiaba","MT"],["campo grande","MS"]];

/* ---------- modelo de publisher novo ---------- */
const FORMATOS_PADRAO = () => ([
  {id:"display", nome:"Display padrão", cpm:12, mix:55, impPorPv:2, ativo:true, chaves:["300x250","728x90","970x250","320x50","320x100","topo","lateral","retangulo","banner","display","billboard"]},
  {id:"diferenciado", nome:"Formatos diferenciados", cpm:22, mix:35, impPorPv:0.6, ativo:true, chaves:["interstitial","intersticial","sticky","anchor","ancora","rewarded","vinheta","inread","in-read","diferenciado","takeover"]},
  {id:"video", nome:"Vídeo", cpm:30, mix:10, impPorPv:0.2, ativo:true, chaves:["video","outstream","instream","preroll","pre-roll"]}
]);
/* ---------- identidade visual por publisher ---------- */
const MARCA_PADRAO = () => ({cor:"#2E7D4F", logoUrl:"", logoAsset:"", logoArquivo:""});
const hexOk = h => /^#[0-9a-f]{6}$/i.test(String(h||"")) ? String(h).toUpperCase() : null;
const rgbDe = h => [1,3,5].map(i=>parseInt(h.slice(i,i+2),16));
const hexDe = a => "#" + a.map(v=>Math.round(Math.max(0,Math.min(255,v))).toString(16).padStart(2,"0")).join("").toUpperCase();
const misturar = (a, b, t) => { const x = rgbDe(a), y = rgbDe(b); return hexDe(x.map((v,i)=>v+(y[i]-v)*t)); };
const lumin = h => { const c = rgbDe(h).map(v=>{ v/=255; return v<=0.03928 ? v/12.92 : Math.pow((v+0.055)/1.055, 2.4); }); return 0.2126*c[0]+0.7152*c[1]+0.0722*c[2]; };
const contraste = (a, b) => { const x = lumin(a), y = lumin(b); return (Math.max(x,y)+0.05)/(Math.min(x,y)+0.05); };
function paleta(cor){
  const c = hexOk(cor) || "#5A2ED8";
  let forte = c; for(let t=0.05; t<=0.95 && contraste(forte,"#FFFFFF")<4.6; t+=0.05) forte = misturar(c,"#000000",t);
  return {accent:c, strong:forte, soft:misturar(c,"#FFFFFF",0.86), sunken:misturar(c,"#FFFFFF",0.94), line:misturar(misturar(c,"#FFFFFF",0.8),"#DADADA",0.55),
    ink: contraste(c,"#1A1A1A") >= contraste(c,"#FFFFFF") ? "#1A1A1A" : "#FFFFFF", muted:misturar(forte,"#5A5A5A",0.72)};
}
function logoUrl(m){
  m = m || {};
  if(/^https:\/\/[^\s"'<>]+$/.test(String(m.logoUrl||""))) return m.logoUrl;
  if(/^[A-Za-z0-9_\/-]+(\.[A-Za-z0-9_-]+)*\.(png|jpe?g|webp|svg)$/.test(String(m.logoArquivo||"")) && !String(m.logoArquivo).includes("..")) return m.logoArquivo;
  return "";
}
let marcaAtual = "";
function aplicarMarca(pub){
  const m = {...MARCA_PADRAO(), ...((pub && pub.marca) || {})}, pl = paleta(m.cor), url = logoUrl(m);
  const chave = JSON.stringify([pl.accent, url, pub && pub.nome]); if(chave===marcaAtual) return; marcaAtual = chave;
  const r = document.documentElement.style;
  r.setProperty("--accent", pl.accent); r.setProperty("--accent-strong", pl.strong); r.setProperty("--accent-soft", pl.soft);
  r.setProperty("--accent-ink", pl.ink); r.setProperty("--sunken", pl.sunken); r.setProperty("--line", pl.line); r.setProperty("--muted", pl.muted);
  const img = $("#brand-logo"), txt = $("#brand-text");
  if(url){ img.src = url; img.alt = (pub && pub.nome) || ""; img.hidden = false; txt.hidden = true; img.onerror = () => { img.hidden = true; txt.hidden = false; }; }
  else { img.hidden = true; txt.hidden = false; }
  txt.textContent = (pub && pub.nome) || "SmarTree";
}
const AUTO_PADRAO = () => ({modo:"email", relatorio:"", planilhaUrl:"", aba:"inventario_planejador", frequencia:"semanal", hora:8});
const novoPublisher = (nome, dominio) => ({marca:MARCA_PADRAO(), auto:AUTO_PADRAO(), nome, dominio:dominio||"", networkCode:"", contato:"", deducao:20, tetoOcupacao:30, premioGeo:0, premioSeg:0, chaveSegmento:"", segmentos:[], cpmValidado:false, pageviewsMes:0, formatos:FORMATOS_PADRAO(), ignorar:[], inventario:null, criadoEm:new Date().toISOString()});

/* ---------- estado ---------- */
const S = {
  db:null, user:null, dl:null, canEdit:false, canWrite:true, dbOff:false, loading:true, loadSlow:false,
  publishers:{}, planos:[], pubId:lsGet("pp:pub"), view:"planejar",
  role:"carregando", me:null, perfil:null, perfis:[], convites:[], confirmConv:null, room:null, assets:null, acessos:[], peers:[], preview:false, uMsg:"", confirmRm:null, pedido:null,
  ufs:new Set(), segs:new Set(), modoSeg:"todos", fmtsSel:null, mergeSel:new Set(), mix:{}, cpm:{}, filtro:"todos", busca:"", confirmDel:null,
  draft:null, draftFor:null, dirty:false, imp:null, creating:false, pubMsg:"", saveMsg:""
};
const curPub = () => S.publishers[S.pubId] || null;
const papel = () => S.role==="admin" && S.preview ? "publisher" : S.role;
const ehAdmin = () => papel()==="admin";
const CAMPOS_PUBLICOS = ["nome","dominio","deducao","tetoOcupacao","premioGeo","premioSeg","formatos","segmentos","inventario","marca","cpmValidado","pageviewsMes"];
const copiaPublica = p => { const o = {}; for(const k of CAMPOS_PUBLICOS) if(p && p[k]!==undefined) o[k] = clone(p[k]); return o; };
/* propostas de usuários de publisher: guardadas no navegador (acesso de leitura não grava na base) */
const chaveLocal = () => `pp:planos:${S.pubId||"-"}:${(S.me&&S.me.id)||"anon"}`;
function lerLocais(){ try{ const v = JSON.parse(lsGet(chaveLocal())||"[]"); return Array.isArray(v) ? v : []; }catch(e){ return []; } }
function gravarLocais(lista){ try{ localStorage.setItem(chaveLocal(), JSON.stringify(lista)); return true; }catch(e){ return false; } }
const salvaLocal = () => false;

/* ---------- inventário ---------- */
function inventario(pub){
  const ativos = (pub.formatos||[]).filter(f=>f.ativo!==false);
  if(pub.inventario && pub.inventario.fonte==="gam" && pub.inventario.porFormatoUF) return {fonte:"gam", dia:pub.inventario.porFormatoUF, cubo:pub.inventario.porFormatoUFSeg||null};
  if(+pub.pageviewsMes > 0){
    const pvDia = pub.pageviewsMes/30, dia = {};
    for(const f of ativos){ dia[f.id] = {}; for(const u of UFS) dia[f.id][u[0]] = pvDia*(+f.impPorPv||0)*(u[2]/POP_TOTAL); }
    return {fonte:"estimado", dia};
  }
  return {fonte:"nenhum", dia:null};
}
const somaObj = o => Object.values(o||{}).reduce((a,b)=>a+(+b||0),0);

function pesosUF(inv, ufs, fmts){
  let w = {}, tot = 0, base = "pop";
  if(inv.dia){
    for(const uf of ufs){ w[uf] = fmts.reduce((a,f)=>a+(+(inv.dia[f.id]||{})[uf]||0),0); tot += w[uf]; }
    base = inv.fonte;
  }
  if(!tot){ w = {}; tot = 0; base = "pop"; for(const uf of ufs){ w[uf] = UF[uf].pop; tot += w[uf]; } }
  for(const uf of ufs) w[uf] = w[uf]/tot;
  return {w, base};
}

function lerForm(){
  const inicio = $("#f-inicio").value, fim = $("#f-fim").value;
  let dias = 0;
  if(inicio && fim){ dias = Math.round((new Date(fim+"T12:00") - new Date(inicio+"T12:00"))/864e5) + 1; }
  return {
    cliente: $("#f-cliente").value.trim(), inicio, fim, dias,
    verba: parseMoney($("#f-verba").value), contexto: $("#f-contexto").value.trim(),
    modoPraca: $("#p-estados").checked ? "estados" : "nacional",
    ufs: UFS.map(u=>u[0]).filter(u=>S.ufs.has(u)), mix:{...S.mix}, cpm:{...S.cpm},
    fmtsSel: curPub() ? fmtsSelecionados(curPub()) : [],
    modoSeg: S.modoSeg, segs: ((curPub()||{}).segmentos||[]).map(x=>String(x.id)).filter(id=>S.segs.has(id)),
    desconto: Math.min(100, Math.max(0, parseFloat(String($("#f-desc").value).replace(",","."))||0))
  };
}

function invDiaDe(inv, fid, p, s){
  if(!inv.dia) return null;
  if(!(fid in inv.dia)) return null;
  if(s==="*") return p==="BR" ? somaObj(inv.dia[fid]) : (+(inv.dia[fid]||{})[p]||0);
  if(!inv.cubo) return null;
  const o = inv.cubo[fid]||{};
  if(p==="BR"){ let t = 0; for(const uf in o) t += +((o[uf]||{})[s])||0; return t; }
  return +((o[p]||{})[s])||0;
}
const segNome = (pub, id) => { const x = (pub.segmentos||[]).find(s=>String(s.id)===String(id)); return x ? `${x.id} · ${x.nome}` : (id==="ND" ? "Sem segmento" : String(id)); };
function calcular(pub, f){
  if(!pub) return {erro:"Escolha um publisher."};
  if(!f.inicio || !f.fim) return {erro:"Informe o início e o fim da campanha."};
  if(f.dias < 1) return {erro:"O fim da campanha precisa ser depois do início."};
  if(!(f.verba > 0)) return {erro:"Informe a verba bruta."};
  const geo = f.modoPraca === "estados";
  if(geo && !f.ufs.length) return {erro:"Escolha ao menos um estado, ou planeje como Nacional."};
  const porSeg = f.modoSeg==="escolher" && (pub.segmentos||[]).length > 0;
  if(porSeg && !f.segs.length) return {erro:"Escolha ao menos um segmento, ou planeje para todos os segmentos."};
  const sel = Array.isArray(f.fmtsSel) ? new Set(f.fmtsSel) : null;
  const fmts = (pub.formatos||[]).filter(x=>x.ativo!==false && (!sel || sel.has(x.id)));
  if(!fmts.length) return {erro:"Escolha ao menos um formato."};
  /* formato escolhido com % padrão zero ainda recebe uma parte mínima; o limite de inventário redistribui o excesso */
  const mixv = fmts.map(x => Math.max(1, +(f.mix[x.id] ?? x.mix) || 0));
  const msum = mixv.reduce((a,b)=>a+b,0);
  if(!msum) return {erro:"Distribua a verba entre os formatos."};
  const ded = (+pub.deducao||0)/100, teto = (+pub.tetoOcupacao||0)/100;
  const premio = (1 + (geo ? (+pub.premioGeo||0)/100 : 0)) * (1 + (porSeg ? (+pub.premioSeg||0)/100 : 0)) - 1;
  const desc = (+f.desconto||0)/100;
  if(desc >= 1) return {erro:"O desconto precisa ser menor que 100%."};
  const inv = inventario(pub);
  const pracas = geo ? f.ufs : ["BR"];
  const segs = porSeg ? f.segs : ["*"];
  const grupos = [];
  for(const p of pracas) for(const s of segs) grupos.push({p, s});
  let pesos = grupos.map(g => { let v = 0; for(const x of fmts){ const d = invDiaDe(inv, x.id, g.p, g.s); if(d==null) return null; v += d; } return v; });
  let base = inv.fonte;
  if(pesos.some(v=>v==null) || !pesos.reduce((a,b)=>a+(b||0),0)){ base = "pop"; pesos = grupos.map(g => (g.p==="BR" ? 1 : UF[g.p].pop)/segs.length); }
  const ptot = pesos.reduce((a,b)=>a+b,0); pesos = pesos.map(v=>v/ptot);
  /* 1) linhas com a verba desejada (peso do grupo × % do formato) */
  const linhas = [];
  grupos.forEach((g, gi) => fmts.forEach((x, i) => {
    const share = mixv[i]/msum; if(!share) return;
    const cpmTab = (+(f.cpm[x.id] ?? x.cpm) || 0)*(1+premio);
    const cpm = cpmTab*(1-desc);
    const d = invDiaDe(inv, x.id, g.p, g.s);
    const invPer = d==null ? null : d*f.dias;
    linhas.push({gi, praca:g.p, seg:g.s, fmt:x, cpmTab, cpm, invPer, desejo: f.verba*pesos[gi]*share, bruto:0});
  }));
  /* 2) limite entregável: inventário do período × teto de ocupação, convertido em verba bruta */
  const fatorTeto = teto > 0 ? teto : 1;
  for(const l of linhas) l.capBruto = (l.invPer==null || !(l.cpm>0)) ? Infinity : l.invPer*fatorTeto*l.cpm/1000/(1-ded);
  /* 3) distribui a verba; o que não cabe numa linha vai para as que têm espaço, na mesma proporção */
  let restante = f.verba, abertas = linhas.filter(l => l.desejo>0 && l.capBruto>0);
  for(let it=0; it<60 && restante>0.005 && abertas.length; it++){
    const wsum = abertas.reduce((a,l)=>a+l.desejo,0); let usado = 0;
    for(const l of abertas){ const add = restante*l.desejo/wsum, cabe = l.capBruto - l.bruto; const v = Math.min(add, cabe); l.bruto += v; usado += v; }
    restante -= usado;
    abertas = abertas.filter(l => l.capBruto - l.bruto > 0.005);
  }
  const naoAlocado = restante > 0.01 ? restante : 0;
  for(const l of linhas){
    l.liq = l.bruto*(1-ded);
    l.imps = l.cpm > 0 ? l.liq/l.cpm*1000 : 0;
    l.ocup = l.invPer > 0 ? l.imps/l.invPer : null;
    l.limitado = l.capBruto !== Infinity && l.capBruto - l.bruto < 0.01;
    l.semInv = l.invPer === 0;
    l.alerta = false;
  }
  const tot = linhas.reduce((a,l)=>({bruto:a.bruto+l.bruto, liq:a.liq+l.liq, imps:a.imps+l.imps, inv:(a.inv==null||l.invPer==null)?null:a.inv+l.invPer}), {bruto:0,liq:0,imps:0,inv:inv.dia?0:null});
  tot.ocup = tot.inv > 0 ? tot.imps/tot.inv : null;
  const limitadas = linhas.filter(l=>l.limitado).length;
  return {naoAlocado, limitadas, fatorTeto, linhas, tot, ded, teto, premio, desc, inv, base, geo, porSeg, pracas, segs, grupos, fmts, mixv, msum};
}

/* ---------- formulário ---------- */
function proxSegunda(){ const d = new Date(); d.setDate(d.getDate() + ((8 - d.getDay()) % 7 || 7)); return d; }
const iso = d => `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`;
function initForm(){
  const ini = proxSegunda(), fim = new Date(ini); fim.setDate(fim.getDate()+29);
  $("#f-cliente").value = "Exemplo — Maquininha Pague Fácil";
  $("#f-inicio").value = iso(ini); $("#f-fim").value = iso(fim);
  $("#f-verba").value = BRL.format(50000);
  $("#f-contexto").value = "Lançamento de maquininha sem mensalidade para pequenos lojistas de São Paulo e do Rio de Janeiro. Tom prático e direto, falando de controle do caixa.";
  $("#regions").innerHTML = Object.entries(REGIOES).map(([k,n]) => `<button type="button" class="btn sm" data-reg="${k}">${n}</button>`).join("") + `<button type="button" class="btn sm" data-reg="limpar">Limpar</button>`;
  $("#ufgrid").innerHTML = UFS.map(u => `<button type="button" data-uf="${u[0]}" aria-pressed="false" title="${esc(u[1])}">${u[0]}</button>`).join("");
  atualizarHints();
}
function setUFs(list){ S.ufs = new Set(list); document.querySelectorAll("#ufgrid button").forEach(b => b.setAttribute("aria-pressed", S.ufs.has(b.dataset.uf) ? "true" : "false")); }
function atualizarHints(){
  const f = lerForm();
  $("#h-dias").textContent = f.dias > 0 ? `${f.dias} dias de campanha` : "";
  const pub = curPub();
  const ded = pub ? (+pub.deducao||0) : 20;
  $("#h-verba").textContent = f.verba > 0 ? `Líquido: ${BRL.format(f.verba*(1-ded/100))} (dedução de ${ded}% de agência e custos)` : `O líquido desconta ${ded}% de agência e custos.`;
  const txt = $("#f-contexto").value;
  const frases = (txt.match(/[^.!?]+[.!?]+|[^.!?]+$/g)||[]).filter(s=>s.trim()).length;
  $("#h-frases").textContent = `${frases} de 2 frases`; $("#h-frases").style.color = frases>2 ? "var(--bad)" : "";
  $("#h-chars").textContent = `${txt.length}/320`;
  $("#uf-box").hidden = !$("#p-estados").checked;
  const n = S.ufs.size;
  const premio = pub ? (+pub.premioGeo||0) : 0;
  $("#h-ufs").textContent = (n ? `${n} estado${n>1?"s":""} selecionado${n>1?"s":""}.` : "Nenhum estado selecionado.") + (premio ? ` CPM com acréscimo de ${premio}% por segmentação geográfica.` : "");
  const fmtsAtivos = pub ? (pub.formatos||[]).filter(x=>x.ativo!==false) : [];
  const d = f.desconto/100;
  if(!fmtsAtivos.length){ $("#h-desc").textContent = "Aplicado sobre o CPM tabela de todos os formatos."; }
  else {
    const tabs = fmtsAtivos.map(x=>+x.cpm||0), mn = Math.min(...tabs), mx = Math.max(...tabs);
    const faixa = (a,b) => a===b ? BRL.format(a) : `${BRL.format(a)} a ${BRL.format(b)}`;
    $("#h-desc").textContent = `Aplicado igualmente aos ${fmtsAtivos.length} formatos. CPM tabela ${faixa(mn,mx)}` + (d > 0 ? ` → negociado ${faixa(mn*(1-d), mx*(1-d))}.` : ".") + (pub.cpmValidado ? "" : " CPMs tabela ainda a validar na configuração.");
  }
  $("#h-desc").style.color = f.desconto >= 100 ? "var(--bad)" : "";
  sugerirPracas();
  sugerirSegs();
  bannerAprovado();
}
function detectarUFs(texto){
  const t = " " + norm(texto) + " ", achou = new Set();
  for(const [nome, uf] of Object.entries(UF_POR_NOME)){
    if(nome==="para") continue; /* "para" é preposição; Pará só com "estado do" ou "no" */
    const re = new RegExp("(^|[^a-z])" + nome.replace(/ /g,"\\s+") + "([^a-z]|$)"); if(re.test(t)) achou.add(uf);
  }
  if(/(^|[^a-z])(estado do|no) para([^a-z]|$)/.test(t)) achou.add("PA");
  if(achou.has("MS") && !/mato grosso(?! do sul)/.test(t)) achou.delete("MT");
  for(const [c, uf] of CIDADES) if(t.includes(" "+c) ) achou.add(uf);
  if(/\brio\b(?! grande)(?! de janeiro)/.test(t) && !achou.size) achou.add("RJ");
  const regs = {SE:/sudeste/, S:/regiao sul|sul do pais|\bno sul\b/, NE:/nordeste/, CO:/centro.?oeste/, N:/regiao norte|\bno norte\b/};
  for(const [r, re] of Object.entries(regs)) if(re.test(t)) UFS.filter(u=>u[3]===r).forEach(u=>achou.add(u[0]));
  const up = String(texto).match(/\b[A-Z]{2}\b/g) || [];
  for(const s of up) if(UF[s] && !["SE","ES","AL","PA","TO","MA"].includes(s)) achou.add(s);
  return UFS.map(u=>u[0]).filter(u=>achou.has(u));
}
function sugerirPracas(){
  const box = $("#ctx-suggest");
  const ufs = detectarUFs($("#f-contexto").value);
  const atuais = [...S.ufs].sort().join(","), novas = [...ufs].sort().join(",");
  if(!ufs.length || ($("#p-estados").checked && atuais === novas)){ box.hidden = true; return; }
  box.hidden = false;
  box.innerHTML = `<div class="banner info"><span>O contexto cita <b>${esc(ufs.join(", "))}</b>.</span><span><button type="button" class="btn sm" id="btn-usar-ufs">Planejar nessas praças</button></span></div>`;
  $("#btn-usar-ufs").onclick = () => { $("#p-estados").checked = true; setUFs(ufs); atualizarHints(); render(); };
}
function aprovadoDoCliente(){
  const k = norm($("#f-cliente").value); if(!k) return null;
  return S.planos.filter(p => p.publisherId===S.pubId && p.status==="aprovado" && norm(p.cliente)===k)
    .sort((a,b)=> String(b.aprovadoEm||b.criadoEm).localeCompare(String(a.aprovadoEm||a.criadoEm)))[0] || null;
}
function bannerAprovado(){
  const box = $("#approved-banner"), p = aprovadoDoCliente();
  if(!p){ box.hidden = true; return; }
  box.hidden = false;
  box.innerHTML = `<div class="banner"><span><b>Proposta aprovada</b> em ${esc(fmtData(p.aprovadoEm||p.criadoEm))}: ${esc(BRL.format(p.verba))}, ${esc(p.modoPraca==="estados" ? (p.ufs||[]).join(", ") : "Nacional")}, desconto de ${esc(String(p.desconto||0).replace(".",","))}%.</span><span><button type="button" class="btn sm" id="btn-usar-aprov">Usar como ponto de partida</button></span></div>`;
  $("#btn-usar-aprov").onclick = () => carregarPlano(p, false);
}
function volSegDia(pub){
  const c = pub && pub.inventario && pub.inventario.porFormatoUFSeg; if(!c) return null;
  const ativos = new Set((pub.formatos||[]).filter(f=>f.ativo!==false).map(f=>f.id)), t = {};
  for(const fid in c){ if(!ativos.has(fid)) continue; for(const uf in c[fid]) for(const s in c[fid][uf]) t[s] = (t[s]||0) + (+c[fid][uf][s]||0); }
  return t;
}
function renderSegs(){
  const pub = curPub(), segs = pub ? (pub.segmentos||[]) : [];
  $("#seg-step").hidden = !segs.length;
  $("#n-desc").textContent = segs.length ? "7" : "6"; $("#n-fsel").textContent = segs.length ? "8" : "7";
  $("#n-desc").textContent = segs.length ? "7" : "6"; $("#n-fsel").textContent = segs.length ? "8" : "7";
  if(!segs.length) return;
  const vol = volSegDia(pub), tot = vol ? somaObj(vol) : 0, pick = S.modoSeg==="escolher";
  $("#segbox").innerHTML = `<div class="seg" role="radiogroup" aria-label="Segmentos"><input type="radio" name="segmode" id="s-todos" value="todos" ${pick?"":"checked"}><label for="s-todos">Todos</label><input type="radio" name="segmode" id="s-escolher" value="escolher" ${pick?"checked":""}><label for="s-escolher">Escolher segmentos</label></div>
    <div class="seglist" ${pick?"":"hidden"}>${segs.map(s => `<label><input type="checkbox" id="sg-${esc(s.id)}" data-seg="${esc(s.id)}" ${S.segs.has(String(s.id))?"checked":""}><span>${esc(s.id)} · ${esc(s.nome)}</span><span class="vol">${vol ? NUM.format(vol[s.id]||0)+"/dia" : ""}</span></label>`).join("")}</div>
    <div class="hint" id="h-segs" style="margin-top:6px"></div>`;
  atualizarHintSegs();
}
function atualizarHintSegs(){
  const el = $("#h-segs"); if(!el) return;
  const pub = curPub(), vol = volSegDia(pub), tot = vol ? somaObj(vol) : 0;
  const n = S.segs.size, pick = S.modoSeg==="escolher";
  let t = pick ? (n ? `${n} segmento${n>1?"s":""} selecionado${n>1?"s":""}${vol && tot ? `, ${pct(([...S.segs].reduce((a,s)=>a+(vol[s]||0),0))/tot)} das impressões` : ""}.` : "Marque um ou mais segmentos.") : "A proposta usa o inventário de todos os segmentos.";
  t += vol ? ` Volume: média diária do GAM.${vol.ND && tot ? ` ${pct(vol.ND/tot)} das impressões sem segmento.` : ""}` : " O volume por segmento aparece depois da importação do GAM com a chave do segmento.";
  if(pick && pub && +pub.premioSeg) t += ` CPM com acréscimo de ${pub.premioSeg}% por segmento.`;
  el.textContent = t;
}
function detectarSegs(texto, pub){
  const t = " " + norm(texto) + " ";
  return (pub.segmentos||[]).filter(s => (s.palavras||[]).some(k => { const n = norm(k); return n && new RegExp("(^|[^a-z])" + n.replace(/[.*+?^${}()|[\]\\]/g,"\\$&").replace(/ /g,"\\s+") + "([^a-z]|$)").test(t); })).map(s=>String(s.id));
}
function sugerirSegs(){
  const box = $("#seg-suggest"), pub = curPub();
  if(!box || !pub || !(pub.segmentos||[]).length){ if(box) box.hidden = true; return; }
  const ids = detectarSegs($("#f-contexto").value, pub);
  if(!ids.length || (S.modoSeg==="escolher" && [...S.segs].sort().join()===[...ids].sort().join())){ box.hidden = true; return; }
  box.hidden = false;
  box.innerHTML = `<div class="banner info"><span>O contexto combina com: <b>${esc(ids.map(id=>segNome(pub,id)).join("; "))}</b>.</span><span><button type="button" class="btn sm" id="btn-usar-segs">Planejar nesses segmentos</button></span></div>`;
  $("#btn-usar-segs").onclick = () => { S.modoSeg = "escolher"; S.segs = new Set(ids); renderSegs(); atualizarHints(); renderResult(); };
}
function fmtsSelecionados(pub){
  const ativos = (pub.formatos||[]).filter(f=>f.ativo!==false).map(f=>f.id);
  return S.fmtsSel ? ativos.filter(id=>S.fmtsSel.has(id)) : ativos;
}
function renderFmtSel(){
  const box = $("#fselbox"), pub = curPub();
  if(!pub){ box.innerHTML = `<p class="hint">Escolha um publisher.</p>`; return; }
  const fmts = (pub.formatos||[]).filter(f=>f.ativo!==false);
  if(!fmts.length){ box.innerHTML = `<p class="hint">Este publisher ainda não tem formatos ativos. Cadastre em Configurar publisher.</p>`; return; }
  const sel = new Set(fmtsSelecionados(pub));
  const inv = pub.inventario && pub.inventario.fonte==="gam" ? (pub.inventario.porFormatoUF||{}) : null;
  box.innerHTML = `<div style="display:flex;gap:6px;flex-wrap:wrap"><button type="button" class="btn sm" data-fall="todos">Todos</button><button type="button" class="btn sm" data-fall="limpar">Limpar</button></div>
    <div class="seglist">${fmts.map(f => `<label><input type="checkbox" id="fs-${esc(f.id)}" data-fsel="${esc(f.id)}" ${sel.has(f.id)?"checked":""}><span>${esc(f.nome)} <span class="note">· ${BRL.format(+f.cpm||0)}</span></span><span class="vol">${inv && (f.id in inv) ? NUM.format(somaObj(inv[f.id]))+"/dia" : ""}</span></label>`).join("")}</div>
    <div class="hint" id="h-fsel" style="margin-top:6px"></div>`;
  atualizarHintFsel();
}
function atualizarHintFsel(){
  const el = $("#h-fsel"), pub = curPub(); if(!el || !pub) return;
  const n = fmtsSelecionados(pub).length, tot = (pub.formatos||[]).filter(f=>f.ativo!==false).length;
  el.textContent = n ? `${n} de ${tot} formato${tot>1?"s":""}. A verba é dividida pelo % padrão de cada formato (Configurar publisher); CPM tabela ao lado do nome.` : "Marque ao menos um formato.";
}
function renderMix(){
  renderSegs();
  renderFmtSel();
  const pub = curPub(), box = $("#mix");
  if(!box) return;
  if(!pub){ box.innerHTML = `<p class="hint">Escolha um publisher.</p>`; return; }
  const fmts = (pub.formatos||[]).filter(f=>f.ativo!==false);
  if(!fmts.length){ box.innerHTML = `<p class="hint">Este publisher ainda não tem formatos. Cadastre em Configurar publisher.</p>`; return; }
  box.innerHTML = `<div class="mrow head"><span>Formato</span><span style="text-align:right">% verba</span><span style="text-align:right">CPM tabela</span></div>` +
    fmts.map(f => `<div class="mrow"><span class="name">${esc(f.nome)}</span>
      <input type="number" min="0" max="100" step="5" id="mix-${esc(f.id)}" data-mix="${esc(f.id)}" value="${esc(S.mix[f.id] ?? f.mix)}" aria-label="Percentual da verba em ${esc(f.nome)}">
      <input type="number" min="0" step="0.5" id="cpm-${esc(f.id)}" data-cpm="${esc(f.id)}" value="${esc(S.cpm[f.id] ?? f.cpm)}" aria-label="CPM tabela de ${esc(f.nome)}"></div>`).join("");
  atualizarHintMix();
}
function atualizarHintMix(){
  if(!$("#h-mix")) return;
  const pub = curPub(); if(!pub) return;
  const fmts = (pub.formatos||[]).filter(f=>f.ativo!==false);
  const soma = fmts.reduce((a,f)=>a+(+(S.mix[f.id] ?? f.mix)||0),0);
  $("#h-mix").innerHTML = (soma && Math.round(soma)!==100 ? `Os percentuais somam ${NUM.format(soma)}%; a verba é dividida na proporção. ` : "") + (pub.cpmValidado ? "" : `<span class="chip warn">CPMs padrão a validar</span>`);
}

/* ---------- proposta ---------- */
const nomePraca = p => p==="BR" ? "Nacional" : `${p} · ${UF[p].nome}`;
function fonteTexto(r, pub){
  if(r.inv.fonte==="gam"){ const i = pub.inventario; return `<span class="chip ok">Inventário real do GAM</span> <span class="note">importado em ${esc(fmtData(i.importadoEm))}, média de ${esc(i.dias)} dias</span>`; }
  if(r.inv.fonte==="estimado") return `<span class="chip warn">Inventário estimado</span> <span class="note">pageviews informados × impressões por pageview; importe o relatório do GAM para usar dados reais</span>`;
  return `<span class="chip">Sem inventário</span> <span class="note">a ocupação aparece depois da importação do GAM${r.geo ? "; praças divididas pela população (IBGE 2022)" : ""}</span>`;
}
function ocupCell(l, r){
  if(l.ocup==null) return l.semInv ? `<span class="chip bad">sem inventário</span>` : `<span class="note">—</span>`;
  const w = Math.min(100, l.ocup/r.fatorTeto*100);
  return `<span class="meter"><span class="bar" title="Parte do limite entregável usada"><i class="${l.limitado?"hi":""}" style="width:${w}%"></i></span>${pct(l.ocup)}${l.limitado?` <span class="chip warn" title="Usa todo o inventário disponível para a campanha (${pct(r.fatorTeto)} do inventário)">no limite</span>`:""}</span>`;
}
function render(){
  const pr = papel(), gate = !["admin","publisher"].includes(pr) || !!S.trocaSenha;
  aplicarMarca(gate ? (S.marcaPublica || null) : (S.view==="publisher" && S.draft && S.draftFor===S.pubId ? S.draft : curPub()));
  renderHeader();
  const vistas = pr==="admin" ? ["planejar","salvos","publisher","usuarios"] : ["planejar","salvos"];
  if(!gate && !vistas.includes(S.view)) S.view = "planejar";
  for(const v of ["planejar","salvos","publisher","usuarios"]){ $("#view-"+v).hidden = gate || S.view!==v; const tb = $("#tab-"+v); tb.hidden = !vistas.includes(v); tb.setAttribute("aria-selected", S.view===v ? "true":"false"); }
  $("#view-acesso").hidden = !gate; $("#tabs").hidden = gate;
  if(gate){ renderAcesso(); return; }
  if(S.view==="planejar") renderResult();
  if(S.view==="salvos") renderSalvos();
  if(S.view==="publisher") renderPublisher();
  if(S.view==="usuarios") renderUsuarios();
}
function renderHeader(){
  const pr = papel(), pub = curPub();
  if(S.avisoSenha){ const m = $("#sysmsg"); setTimeout(() => { if(S.avisoSenha){ m.hidden = false; m.textContent = S.avisoSenha; S.avisoSenha = ""; setTimeout(()=>{ m.hidden = true; }, 6000); } }, 0); }
  $("#pubpick").hidden = pr!=="admin";
  $("#pub-static").hidden = pr!=="publisher" || !pub; if(pub) $("#pub-static").textContent = pub.nome;
  $("#brand-sub").textContent = pr==="publisher" && pub ? `Propostas comerciais · ${pub.nome}` : "Formatos, praças e segmentos · inventário do GAM";
  $("#previewbar").hidden = !(S.role==="admin" && S.preview);
  if(S.role==="admin" && S.preview) $("#preview-txt").textContent = `Visualizando como um usuário de ${pub ? pub.nome : "publisher"} vê o planejador.`;
  const uc = $("#userchip");
  if(S.me && S.me.id){
    const rot = S.role==="admin" ? "Admin" : S.role==="publisher" ? "Publisher" : "Sem acesso";
    uc.innerHTML = `<img alt="" id="uc-img" title="${esc((S.me.name||"Você") + " · " + rot)}"><div><b id="uc-nome"></b><span>${esc(rot)}</span></div>${S.role==="admin" && !S.preview && pub ? `<button class="btn sm" id="btn-preview" title="Ver o planejador como um usuário deste publisher">Ver como publisher</button>` : ""}<button class="btn sm" id="btn-senha" title="Definir ou trocar sua senha">Senha</button><button class="btn sm" id="btn-sair">Sair</button>`;
    $("#uc-img").src = S.me.avatarUrl; $("#uc-nome").textContent = S.me.name || "Você";
  } else uc.innerHTML = "";
  if(pr==="admin"){
    const sel = $("#pub-select");
    const ids = Object.keys(S.publishers).sort((a,b)=>String(S.publishers[a].nome).localeCompare(String(S.publishers[b].nome)));
    let html = ids.map(id => `<option value="${esc(id)}"${id===S.pubId?" selected":""}>${esc(S.publishers[id].nome)}</option>`).join("");
    if(!ids.length) html = `<option value="">${S.loading ? "Carregando…" : "Nenhum publisher"}</option>`;
    if(S.canEdit && !S.dbOff) html += `<option value="__novo">+ Novo publisher</option>`;
    sel.innerHTML = html;
  }
  const n = S.planos.filter(p=>p.publisherId===S.pubId).length;
  $("#saved-count").textContent = n ? String(n) : "";
  const msg = $("#sysmsg");
  if(S.dbOff){ msg.hidden = false; msg.textContent = "Os dados salvos não estão disponíveis nesta visualização. Você pode montar a proposta e exportar, mas salvar e configurar publishers ficam desativados."; }
  else if(S.loading && S.loadSlow){ msg.hidden = false; msg.textContent = "Os dados ainda não carregaram. Verifique a conexão; a página continua tentando."; }
  else msg.hidden = true;
}
function renderAcesso(){
  const box = $("#view-acesso");
  if(S.role==="config"){ box.innerHTML = `<div class="panel gate"><h2>Configuração pendente</h2><p>Preencha o endereço e a chave pública do Supabase no arquivo <b>config.js</b> e publique de novo. O passo a passo está no README do repositório.</p></div>`; return; }
  if(S.role==="carregando"){ box.innerHTML = `<div class="panel gate"><h2>Carregando seu acesso…</h2><p class="note">Identificando sua conta e o publisher liberado para você.</p></div>`; return; }
  if(S.trocaSenha){
    box.innerHTML = `<div class="panel gate">
      <h2>${S.trocaSenha==="recuperar" ? "Crie uma nova senha" : "Definir senha"}</h2>
      <p>${S.me && S.me.email ? `Conta: <b>${esc(S.me.email)}</b>. ` : ""}Use pelo menos 8 caracteres.</p>
      <form class="gate-form" id="senha-form" novalidate>
        <label>Nova senha<input type="password" id="ns-1" autocomplete="new-password" minlength="8" required></label>
        <label>Repita a nova senha<input type="password" id="ns-2" autocomplete="new-password" minlength="8" required></label>
        <div style="display:flex;gap:8px;flex-wrap:wrap"><button class="btn primary" type="submit" id="btn-ns">Salvar senha</button>${S.trocaSenha==="recuperar" ? "" : `<button class="btn" type="button" id="btn-ns-cancel">Cancelar</button>`}</div>
      </form>
      <div class="status" id="ns-msg" aria-live="polite"></div>
    </div>`;
    $("#senha-form").onsubmit = async e => {
      e.preventDefault();
      const a = $("#ns-1").value, b = $("#ns-2").value, msg = $("#ns-msg");
      if(a.length < 8){ msg.className = "status bad"; msg.textContent = "A senha precisa ter pelo menos 8 caracteres."; return; }
      if(a !== b){ msg.className = "status bad"; msg.textContent = "As duas senhas não são iguais."; return; }
      $("#btn-ns").disabled = true; msg.className = "status"; msg.textContent = "Salvando…";
      const err = await DATA.definirSenha(a);
      if(err){ $("#btn-ns").disabled = false; msg.className = "status bad"; msg.textContent = /different|same/i.test(err) ? "Use uma senha diferente da atual." : "Não foi possível salvar a senha agora. Tente de novo."; return; }
      S.trocaSenha = null; S.uMsg = ""; await DATA.carregarSessao(); S.avisoSenha = "Senha salva. Nos próximos acessos, entre com e-mail e senha."; render();
    };
    const c = $("#btn-ns-cancel"); if(c) c.onclick = () => { S.trocaSenha = null; render(); };
    return;
  }
  if(S.role==="login"){
    const modo = S.loginModo || "entrar";
    const titulo = {entrar: S.marcaPublica ? `Entrar no planejador da ${S.marcaPublica.nome}` : "Entrar no planejador", criar:"Criar conta", esqueci:"Recuperar senha"}[modo];
    box.innerHTML = `<div class="panel gate">
      <h2>${titulo}</h2>
      ${modo==="criar" ? `<p>Recebeu um convite? Crie sua conta com o <b>mesmo e-mail do convite</b> e você entra direto no seu publisher.</p>` : modo==="esqueci" ? `<p>Informe seu e-mail. Enviamos um link para você criar uma nova senha.</p>` : ""}
      <form class="gate-form" id="login-form" novalidate>
        <label>E-mail<input type="email" id="login-email" autocomplete="email" required value="${esc(S.loginEmail||"")}"></label>
        ${modo!=="esqueci" ? `<label>Senha<input type="password" id="login-senha" autocomplete="${modo==="criar" ? "new-password" : "current-password"}" required></label>` : ""}
        ${modo==="criar" ? `<label>Repita a senha<input type="password" id="login-senha2" autocomplete="new-password" required></label>` : ""}
        <button class="btn primary" type="submit" id="btn-login">${{entrar:"Entrar", criar:"Criar conta", esqueci:"Enviar link"}[modo]}</button>
      </form>
      <div class="status ${S.loginOk?"ok":""}" id="login-msg" aria-live="polite">${esc(S.loginOk||"")}</div>
      <div class="gate-links">${modo==="entrar" ? `<button class="linkbtn" data-lmodo="criar">Primeiro acesso? Criar conta</button><button class="linkbtn" data-lmodo="esqueci">Esqueci minha senha</button>` : `<button class="linkbtn" data-lmodo="entrar">Já tenho conta: entrar</button>`}</div>
    </div>`;
    document.querySelectorAll("[data-lmodo]").forEach(b => b.onclick = () => { S.loginModo = b.dataset.lmodo; S.loginOk = ""; S.loginEmail = $("#login-email").value.trim(); renderAcesso(); });
    $("#login-form").onsubmit = async e => {
      e.preventDefault();
      const email = $("#login-email").value.trim().toLowerCase(), msg = $("#login-msg"), btn = $("#btn-login");
      const ruim = t => { msg.className = "status bad"; msg.textContent = t; btn.disabled = false; };
      if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return ruim("Informe um e-mail válido.");
      S.loginEmail = email; btn.disabled = true; msg.className = "status"; msg.textContent = "Aguarde…";
      if(modo==="esqueci"){
        const err = await DATA.esqueciSenha(email);
        if(err) return ruim(/rate|seconds|limit/i.test(err) ? "Muitos pedidos seguidos. Aguarde alguns minutos e tente de novo." : "Não foi possível enviar agora. Tente de novo em instantes.");
        S.loginOk = `Se ${email} tiver conta, enviamos o link para criar uma nova senha. Confira também a caixa de spam.`; S.loginModo = "entrar"; renderAcesso(); return;
      }
      const senha = $("#login-senha").value;
      if(modo==="criar"){
        if(senha.length < 8) return ruim("A senha precisa ter pelo menos 8 caracteres.");
        if(senha !== $("#login-senha2").value) return ruim("As duas senhas não são iguais.");
        const r = await DATA.criarConta(email, senha);
        if(r.erro) return ruim(/registered|already/i.test(r.erro) ? "Este e-mail já tem conta. Use Entrar, ou Esqueci minha senha." : /password/i.test(r.erro) ? "Senha fraca. Use pelo menos 8 caracteres, misturando letras e números." : "Não foi possível criar a conta agora. Tente de novo.");
        if(r.confirmar){ S.loginOk = `Conta criada. Enviamos um e-mail de confirmação para ${email}; confirme e depois entre com e-mail e senha.`; S.loginModo = "entrar"; renderAcesso(); }
        return;
      }
      if(!senha) return ruim("Informe sua senha.");
      const err = await DATA.entrar(email, senha);
      if(err) return ruim(/invalid/i.test(err) ? "E-mail ou senha incorretos." : /confirm/i.test(err) ? "Confirme seu e-mail pelo link que enviamos e tente de novo." : "Não foi possível entrar agora. Tente de novo.");
    };
    return;
  }
  if(S.role==="outro-publisher"){
    const url = enderecoDoPublisher(S.pubCerto);
    box.innerHTML = `<div class="panel gate"><h2>Este endereço é de outro publisher</h2>
      <p>Seu acesso é ao planejador de outro publisher. Entre pelo seu endereço:</p>
      <p><a class="btn primary" href="${esc(url)}">${esc(url.replace(/^https:\/\//,"").replace(/\/$/,""))}</a></p>
      <p><button class="btn" id="btn-sair3">Sair</button></p></div>`;
    $("#btn-sair3").onclick = () => DATA.sair();
    return;
  }
  const emp = (S.perfil && S.perfil.empresa) || (S.marcaPublica && S.marcaPublica.nome) || "";
  box.innerHTML = `<div class="panel gate">
    <h2>Você ainda não tem acesso${S.marcaPublica ? ` ao planejador da ${esc(S.marcaPublica.nome)}` : " a um publisher"}</h2>
    <p>Você entrou como <b>${esc((S.me && S.me.email) || "")}</b>. O acesso é liberado pela equipe do planejador. Informe sua empresa para que ela identifique seu pedido.</p>
    <div class="row"><label>Empresa<input type="text" id="acc-emp" maxlength="80" placeholder="Ex.: MarketUP" value="${esc(emp)}"></label>
      <button class="btn primary" id="btn-pedir">${emp ? "Atualizar pedido" : "Pedir acesso"}</button></div>
    <div class="status ${emp?"ok":""}" id="acc-msg">${emp ? "Pedido registrado. Assim que o acesso for liberado, clique em Verificar acesso." : ""}</div>
    <div style="display:flex;gap:8px;flex-wrap:wrap"><button class="btn" id="btn-verificar">Verificar acesso</button><button class="btn" id="btn-sair2">Sair</button></div>
  </div>`;
  $("#btn-pedir").onclick = async () => {
    const v = $("#acc-emp").value.trim().slice(0,80), msg = $("#acc-msg");
    if(!v){ msg.className = "status bad"; msg.textContent = "Informe o nome da sua empresa."; return; }
    const ok = await DATA.pedirAcesso(v);
    if(!ok){ msg.className = "status bad"; msg.textContent = "Não foi possível registrar agora. Tente de novo."; return; }
    S.perfil = {...(S.perfil||{}), empresa:v}; renderAcesso();
  };
  $("#btn-verificar").onclick = () => DATA.carregarSessao();
  $("#btn-sair2").onclick = () => DATA.sair();
}
function renderResult(){
  const box = $("#result"), pub = curPub();
  if(!pub){
    box.innerHTML = S.loading && !S.dbOff ? `<div class="panel empty">Carregando publishers…</div>` :
      `<div class="panel empty"><b>Nenhum publisher cadastrado.</b><span>Cadastre o publisher para montar propostas com os formatos, CPMs e inventário dele.</span>${S.canEdit?`<button class="btn primary" data-go="publisher">Configurar publisher</button>`:""}</div>`;
    return;
  }
  const f = lerForm(), r = calcular(pub, f);
  if(r.erro){ box.innerHTML = `<div class="panel empty">${esc(r.erro)}</div>`; return; }
  const exemplo = /^exemplo/i.test(f.cliente);
  let rows = "";
  const rotulo = g => esc(nomePraca(g.p)) + (g.s!=="*" ? `<span class="segname">${esc(segNome(pub, g.s))}</span>` : "");
  r.grupos.forEach((g, gi) => {
    const ls = r.linhas.filter(l=>l.gi===gi);
    ls.forEach((l,i) => {
      rows += `<tr><td class="praca">${i===0 ? rotulo(g) : ""}</td><td>${esc(l.fmt.nome)}</td><td class="num">${BRL.format(l.cpmTab)}</td><td class="num">${r.desc?pct(r.desc):"—"}</td><td class="num"><b>${BRL.format(l.cpm)}</b></td><td class="num">${NUM.format(l.imps)}</td><td class="num">${l.invPer==null?"—":NUM.format(l.invPer)}</td><td class="num">${ocupCell(l,r)}</td><td class="num">${BRL.format(l.bruto)}</td><td class="num">${BRL.format(l.liq)}</td></tr>`;
    });
    if(r.grupos.length>1){
      const s = ls.reduce((a,l)=>({imps:a.imps+l.imps,bruto:a.bruto+l.bruto,liq:a.liq+l.liq,inv:(a.inv==null||l.invPer==null)?null:a.inv+l.invPer}),{imps:0,bruto:0,liq:0,inv:0});
      const nomeSub = [g.p!=="BR" ? g.p : "", g.s!=="*" ? "segmento " + g.s : ""].filter(Boolean).join(" · ");
      rows += `<tr class="subt"><td colspan="5">Subtotal ${esc(nomeSub)}</td><td class="num">${NUM.format(s.imps)}</td><td class="num">${s.inv==null?"—":NUM.format(s.inv)}</td><td class="num">${s.inv>0?pct(s.imps/s.inv):"—"}</td><td class="num">${BRL.format(s.bruto)}</td><td class="num">${BRL.format(s.liq)}</td></tr>`;
    }
  });
  const alertas = r.linhas.filter(l=>l.alerta).length;
  const pracasTxt = r.geo ? r.pracas.join(", ") : "Nacional";
  const segsTxt = r.porSeg ? r.segs.map(id=>segNome(pub,id)).join("; ") : "Todos";
  const divisao = r.grupos.length<2 ? "" : r.base==="pop" ? (r.geo && r.porSeg ? "Pela população de cada estado e em partes iguais entre os segmentos" : r.geo ? "Pela população de cada estado (IBGE 2022)" : "Em partes iguais entre os segmentos") + (r.porSeg && !r.inv.cubo ? ", porque ainda não há volume por segmento do GAM" : "") : "Pelo inventário de cada " + (r.geo && r.porSeg ? "estado e segmento" : r.geo ? "estado" : "segmento") + " no GAM";
  box.innerHTML = `
    <div class="result-head"><h2>Proposta ${esc(pub.nome)}</h2>${exemplo?`<span class="chip accent">Exemplo</span>`:""}
</div>
    <div class="cards">
      <div class="card"><span class="k">Impressões entregáveis</span><span class="v">${NUM.format(r.tot.imps)}</span><span class="note">${NUM.format(r.tot.imps/f.dias)} por dia${r.tot.inv!=null?` · limite ${pct(r.fatorTeto)} do inventário`:""}</span></div>
      <div class="card"><span class="k">Investimento bruto</span><span class="v">${BRL.format(r.tot.bruto)}</span><span class="note">${r.desc ? `desconto de ${pct(r.desc)} no CPM` : "CPM tabela, sem desconto"}</span></div>
      <div class="card"><span class="k">Investimento líquido</span><span class="v">${BRL.format(r.tot.liq)}</span><span class="note">após ${pct(r.ded)} de agência e custos</span></div>
      <div class="card"><span class="k">Ocupação do inventário</span><span class="v">${r.tot.ocup==null?"—":pct(r.tot.ocup)}</span><span class="note">${r.limitadas?`${r.limitadas} linha${r.limitadas>1?"s":""} no limite do inventário`:`teto por campanha: ${pct(r.fatorTeto)}`}</span></div>
    </div>
    ${r.naoAlocado ? `<div class="banner" style="border-color:var(--warn);background:var(--warn-soft)"><span><b style="color:var(--warn)">O inventário não comporta toda a verba.</b> Com estas praças, segmentos e o teto de ${pct(r.fatorTeto)} por campanha, dá para entregar até ${BRL.format(r.tot.bruto)} de verba bruta. Ficam sem entrega ${BRL.format(r.naoAlocado)}.</span><span style="display:flex;gap:8px;flex-wrap:wrap"><button class="btn sm" id="btn-ajusta-verba">Ajustar verba para ${BRL.format(Math.floor(r.tot.bruto*100)/100)}</button></span></div>` : ""}
    <div>${fonteTexto(r,pub)}</div>
    <div class="tablebox"><table class="plan">
      <tbody class="meta">
        <tr><th>Publisher</th><td colspan="9">${esc(pub.nome)}${pub.dominio?` · ${esc(pub.dominio)}`:""}</td></tr>
        <tr><th>Anunciante</th><td colspan="9">${esc(f.cliente||"—")}</td></tr>
        <tr><th>Período</th><td colspan="9">${esc(fmtData(f.inicio))} a ${esc(fmtData(f.fim))} · ${f.dias} dias</td></tr>
        <tr><th>Verba bruta</th><td colspan="9">${BRL.format(f.verba)}</td></tr>
        <tr><th>Contexto</th><td colspan="9">${esc(f.contexto||"—")}</td></tr>
        <tr><th>Praças</th><td colspan="9">${esc(pracasTxt)}</td></tr>
        ${(pub.segmentos||[]).length ? `<tr><th>Segmentos</th><td colspan="9">${esc(segsTxt)}</td></tr>` : ""}
        ${divisao ? `<tr><th>Divisão da verba</th><td colspan="9">${esc(divisao)}</td></tr>` : ""}
        <tr><th>Desconto</th><td colspan="9">${r.desc ? pct(r.desc)+" sobre o CPM tabela" : "Sem desconto"}</td></tr>
      </tbody>
      <tbody class="cols"><tr><th>${r.porSeg ? "Praça / segmento" : "Praça"}</th><th>Formato</th><th class="num">CPM tabela</th><th class="num">Desconto</th><th class="num">CPM negociado</th><th class="num">Impressões</th><th class="num">Inventário no período</th><th class="num">Ocupação</th><th class="num">Invest. bruto</th><th class="num">Invest. líquido</th></tr></tbody>
      <tbody>${rows}</tbody>
      <tfoot><tr><td colspan="5">Total</td><td class="num">${NUM.format(r.tot.imps)}</td><td class="num">${r.tot.inv==null?"—":NUM.format(r.tot.inv)}</td><td class="num">${r.tot.ocup==null?"—":pct(r.tot.ocup)}</td><td class="num">${BRL.format(r.tot.bruto)}</td><td class="num">${BRL.format(r.tot.liq)}</td></tr>${r.naoAlocado?`<tr><td colspan="8" style="background:var(--warn-soft);color:var(--warn)">Verba sem entrega (não cabe no inventário)</td><td class="num" style="background:var(--warn-soft);color:var(--warn)">${BRL.format(r.naoAlocado)}</td><td class="num" style="background:var(--warn-soft)"></td></tr>`:""}</tfoot>
    </table></div>
    <p class="note">CPM negociado = CPM tabela × (1 − desconto). Impressões = investimento líquido ÷ CPM negociado × 1.000${r.premio?`; CPM tabela com ${pct(r.premio)} de acréscimo pela segmentação (estado e/ou segmento)`:""}. Ocupação = impressões ÷ inventário do formato na praça e segmento no período; cada linha entrega no máximo ${pct(r.fatorTeto)} desse inventário (teto por campanha), e a verba excedente vai para as linhas com espaço.</p>
    <div class="actions" style="margin-left:0;justify-content:flex-end;border-top:1px solid var(--line);padding-top:16px">
      <span class="status" id="save-status" style="margin-right:auto">${esc(S.saveMsg)}</span>
      <button class="btn" id="btn-export" ${S.dl?"":"disabled title=\"Exportação indisponível nesta visualização\""}>Exportar Excel</button>
      <button class="btn" id="btn-ppt" ${S.dl?"":"disabled title=\"Exportação indisponível nesta visualização\""}>Exportar apresentação</button>
      <button class="btn primary" id="btn-salvar" ${salvaLocal() || (S.db && S.canWrite) ? "" : "disabled"} ${salvaLocal() ? 'title="As propostas ficam salvas neste navegador. Exporte em Excel para compartilhar."' : ""}>${salvaLocal() ? "Salvar neste navegador" : "Salvar proposta"}</button>
    </div>`;
  $("#btn-export").onclick = () => exportar(pub, f, r);
  $("#btn-ppt").onclick = async e => { const b = e.currentTarget; b.disabled = true; const o = b.textContent; b.textContent = "Gerando…"; try{ await exportarPPT(pub, f, r); } finally { b.disabled = !S.dl; b.textContent = o; } };
  const aj = $("#btn-ajusta-verba"); if(aj) aj.onclick = () => { $("#f-verba").value = BRL.format(Math.floor(r.tot.bruto*100)/100); S.saveMsg = ""; atualizarHints(); renderResult(); };
  $("#btn-salvar").onclick = () => salvarPlano(pub, f, r);
}

/* ---------- salvar / abrir / aprovar ---------- */
async function salvarPlano(pub, f, r){
  if(!salvaLocal() && !S.db) return;
  if(!f.cliente){ S.saveMsg = ""; const st=$("#save-status"); st.className="status bad"; st.textContent="Informe o anunciante antes de salvar."; return; }
  const btn = $("#btn-salvar"); btn.disabled = true;
  const mix = {}, cpm = {};
  r.fmts.forEach((x,i)=>{ mix[x.id] = r.mixv[i]; cpm[x.id] = +(f.cpm[x.id] ?? x.cpm)||0; });
  const doc = {publisherId:S.pubId, publisherNome:pub.nome, cliente:f.cliente, inicio:f.inicio, fim:f.fim, dias:f.dias, verba:f.verba, contexto:f.contexto,
    modoPraca:f.modoPraca, ufs:f.modoPraca==="estados"?f.ufs:[], mix, cpm, desconto:f.desconto, modoSeg:f.modoSeg, segs:f.modoSeg==="escolher"?f.segs:[], formatosSel:f.fmtsSel, status:"salvo", criadoEm:new Date().toISOString(),
    totais:{imps:Math.round(r.tot.imps), bruto:+r.tot.bruto.toFixed(2), liq:+r.tot.liq.toFixed(2)}, fonteInventario:r.inv.fonte};
  if(salvaLocal()){
    const lista = lerLocais(); lista.push({...doc, id:"l"+Date.now().toString(36)});
    const st = $("#save-status");
    if(gravarLocais(lista)){ S.planos = lista; st.className="status ok"; st.textContent = `Proposta salva neste navegador para ${f.cliente}. Ela aparece em Propostas salvas.`; renderHeader(); bannerAprovado(); }
    else { st.className="status bad"; st.textContent = "Este navegador não permite salvar. Exporte a proposta em Excel."; }
    btn.disabled = false; return;
  }
  try{ await S.db.collection("planos").add(doc); const st=$("#save-status"); st.className="status ok"; st.textContent=`Proposta salva para ${f.cliente}. Ela aparece em Propostas salvas.`; }
  catch(e){ const st=$("#save-status"); st.className="status bad"; st.textContent = e && e.code==="quota_exceeded" ? "O limite de propostas salvas foi atingido. Exclua propostas antigas e tente de novo." : "Não foi possível salvar agora. Tente novamente em instantes."; if(e && e.code==="invalid_argument") S.canWrite=false; }
  finally{ btn.disabled = !(S.db && S.canWrite); }
}
function carregarPlano(p, comPeriodo){
  if(p.publisherId && p.publisherId!==S.pubId && S.publishers[p.publisherId]){ S.pubId = p.publisherId; lsSet("pp:pub", S.pubId); }
  $("#f-cliente").value = p.cliente||"";
  if(comPeriodo){ $("#f-inicio").value = p.inicio||""; $("#f-fim").value = p.fim||""; }
  $("#f-verba").value = BRL.format(p.verba||0);
  $("#f-contexto").value = p.contexto||"";
  const est = p.modoPraca==="estados";
  $("#p-estados").checked = est; $("#p-nacional").checked = !est;
  setUFs(est ? (p.ufs||[]) : []);
  S.mix = {}; S.cpm = {};
  S.fmtsSel = Array.isArray(p.formatosSel) ? new Set(p.formatosSel) : null;
  S.modoSeg = p.modoSeg==="escolher" ? "escolher" : "todos"; S.segs = new Set((p.segs||[]).map(String));
  $("#f-desc").value = p.desconto||0;
  S.view = "planejar"; S.saveMsg = "";
  renderMix(); atualizarHints(); render();
}
async function setStatus(id, status){
  if(salvaLocal()){ const lista = lerLocais().map(p => p.id===id ? {...p, status, aprovadoEm: status==="aprovado" ? new Date().toISOString() : null} : p); gravarLocais(lista); S.planos = lista; renderSalvos(); bannerAprovado(); return; }
  try{ await S.db.collection("planos").doc(id).update(status==="aprovado" ? {status, aprovadoEm:new Date().toISOString()} : {status, aprovadoEm:null}); }
  catch(e){ alertaSalvos("Não foi possível alterar o status agora."); }
}
async function excluir(id){
  if(salvaLocal()){ const lista = lerLocais().filter(p=>p.id!==id); gravarLocais(lista); S.planos = lista; S.confirmDel = null; renderSalvos(); renderHeader(); return; }
  try{ await S.db.collection("planos").doc(id).delete(); } catch(e){ alertaSalvos("Não foi possível excluir agora."); }
  S.confirmDel = null; renderSalvos();
}
function alertaSalvos(t){ const b = $("#saved-box"); const p = document.createElement("p"); p.className="status bad"; p.style.padding="8px 12px"; p.textContent=t; b.prepend(p); }

function renderSalvos(){
  const box = $("#saved-box");
  document.querySelectorAll("[data-filter]").forEach(b => b.classList.toggle("primary", b.dataset.filter===S.filtro));
  if(S.loading && !S.dbOff){ box.innerHTML = `<div class="empty">Carregando propostas…</div>`; return; }
  const k = norm(S.busca);
  const lista = S.planos.filter(p => p.publisherId===S.pubId && (S.filtro==="todos" || p.status===S.filtro) && (!k || norm(p.cliente).includes(k)))
    .sort((a,b)=>String(b.criadoEm).localeCompare(String(a.criadoEm)));
  if(!lista.length){ box.innerHTML = `<div class="empty">${S.planos.some(p=>p.publisherId===S.pubId) ? "Nenhuma proposta com esse filtro." : (salvaLocal() ? "Nenhuma proposta salva neste navegador. Monte uma em Nova proposta e clique em Salvar neste navegador." : "Nenhuma proposta salva para este publisher. Monte uma em Nova proposta e clique em Salvar proposta.")}</div>`; return; }
  box.innerHTML = `<table class="saved"><thead><tr><th>Anunciante</th><th>Período</th><th>Praças / segmentos</th><th class="num">Verba bruta</th><th class="num">Desconto</th><th class="num">Impressões</th><th>Status</th><th class="num">Ações</th></tr></thead><tbody>${
    lista.map(p => `<tr>
      <td><b>${esc(p.cliente)}</b><span class="ctx" title="${esc(p.contexto)}">${esc(p.contexto)}</span></td>
      <td>${esc(fmtData(p.inicio))} a ${esc(fmtData(p.fim))}</td>
      <td>${esc(p.modoPraca==="estados" ? (p.ufs||[]).join(", ") : "Nacional")}${p.modoSeg==="escolher" && (p.segs||[]).length ? `<span class="ctx">Segmentos ${esc(p.segs.join(", "))}</span>` : ""}</td>
      <td class="num">${BRL.format(p.verba||0)}</td>
      <td class="num">${p.desconto ? pct(p.desconto/100) : "—"}</td>
      <td class="num">${NUM.format((p.totais||{}).imps||0)}</td>
      <td>${p.status==="aprovado" ? `<span class="chip ok">Aprovada ${esc(fmtData(p.aprovadoEm))}</span>` : `<span class="chip warn">Aguardando aprovação</span>`}</td>
      <td class="num"><div class="rowact">
        <button class="btn sm" data-abrir="${esc(p.id)}">Abrir</button>
        <button class="btn sm" data-exp="${esc(p.id)}" ${S.dl?"":"disabled"}>Excel</button>
        <button class="btn sm" data-ppt="${esc(p.id)}" ${S.dl?"":"disabled"}>PPT</button>
        ${(S.canWrite || salvaLocal()) ? (p.status==="aprovado" ? `<button class="btn sm" data-st="${esc(p.id)}" data-to="salvo">Desfazer aprovação</button>` : `<button class="btn sm primary" data-st="${esc(p.id)}" data-to="aprovado">Marcar aprovada</button>`) : ""}
        ${(S.canWrite || salvaLocal()) ? (S.confirmDel===p.id ? `<button class="btn sm danger" data-del="${esc(p.id)}" data-confirm="1">Confirmar exclusão</button>` : `<button class="btn sm" data-del="${esc(p.id)}">Excluir</button>`) : ""}
      </div></td></tr>`).join("")}</tbody></table>`;
}

/* ---------- Excel ---------- */
async function exportar(pub, f, r){
  if(!S.dl || !window.ExcelJS) return;
  const pl = paleta((pub.marca||{}).cor), argb = h => "FF" + h.slice(1).toUpperCase();
  const C = {sunken:argb(pl.sunken), line:argb(pl.line), muted:argb(pl.muted), fg:"FF1A1A1A", soft:argb(pl.soft), green:argb(pl.accent), warnSoft:"FFFBF1DF", warn:"FF9A6414"};
  const MONEY = '"R$" #,##0.00', INT = '#,##0', PC = '0.0%', LAST = "J", N = 10;
  const font = (o={}) => ({name:"Calibri", size:11, color:{argb:C.fg}, ...o});
  const fill = argb => ({type:"pattern", pattern:"solid", fgColor:{argb}});
  const linha = {bottom:{style:"thin", color:{argb:C.line}}};
  const wb = new ExcelJS.Workbook(); wb.creator = "Planejador para Publisher";
  const ws = wb.addWorksheet("Proposta", {views:[{showGridLines:false}], pageSetup:{orientation:"landscape", fitToPage:true, fitToWidth:1, fitToHeight:0}});
  ws.columns = [30,22,13,11,15,14,18,12,16,16].map(width => ({width}));

  /* topo: logo, título e faixa verde */
  ws.getRow(1).height = 34;
  try{ const lu = logoUrl(pub.marca); const resp = lu ? await fetch(lu) : null; if(resp && resp.ok){ const u8 = new Uint8Array(await resp.arrayBuffer()); let bin = ""; for(let i=0; i<u8.length; i+=0x8000) bin += String.fromCharCode.apply(null, u8.subarray(i, i+0x8000)); const ext = u8[0]===0x89 && u8[1]===0x50 ? "png" : u8[0]===0xFF && u8[1]===0xD8 ? "jpeg" : null;
      if(ext){ const dims = await new Promise(res => { const im = new Image(); im.onload = () => res([im.naturalWidth, im.naturalHeight]); im.onerror = () => res([124,34]); im.src = lu; });
        const h = 34, w = Math.min(170, Math.round(h * (dims[0]/(dims[1]||1)) || 124));
        const id = wb.addImage({base64:`data:image/${ext};base64,` + btoa(bin), extension:ext}); ws.addImage(id, {tl:{col:0.08, row:0.12}, ext:{width:w, height:h}}); } } }catch(e){}
  ws.mergeCells(`C1:${LAST}1`);
  Object.assign(ws.getCell("C1"), {value:`Proposta comercial · ${pub.nome}`, font:font({size:14, bold:true}), alignment:{horizontal:"right", vertical:"middle"}});
  ws.getRow(2).height = 6;
  for(let c=1; c<=N; c++) ws.getRow(2).getCell(c).border = {bottom:{style:"medium", color:{argb:C.green}}};

  /* cabeçalho do plano */
  const meta = [
    ["Publisher", pub.nome + (pub.dominio ? ` · ${pub.dominio}` : "")],
    ["Anunciante", f.cliente || "—"],
    ["Período", `${fmtData(f.inicio)} a ${fmtData(f.fim)} · ${f.dias} dias`],
    ["Verba bruta", f.verba, MONEY],
    ["Contexto", f.contexto || "—", null, "ctx"],
    ["Praças", r.geo ? r.pracas.join(", ") : "Nacional"]
  ];
  if((pub.segmentos||[]).length) meta.push(["Segmentos", r.porSeg ? r.segs.map(id=>segNome(pub,id)).join("; ") : "Todos"]);
  meta.push(["Desconto", r.desc, '0.0%" sobre o CPM tabela";-0.0%;"Sem desconto"', "desc"]);
  meta.push(["Dedução (agência e custos)", r.ded, '0%', "ded"]);
  let row = 4; const pos = {};
  for(const [label, value, fmt, key] of meta){
    const rr = ws.getRow(row);
    ws.mergeCells(`B${row}:${LAST}${row}`);
    for(let c=1; c<=N; c++){ const cc = rr.getCell(c); cc.fill = fill(C.sunken); cc.border = linha; }
    Object.assign(rr.getCell(1), {value:label.toUpperCase(), font:font({size:9, bold:true, color:{argb:C.muted}}), alignment:{vertical:"middle"}});
    const v = rr.getCell(2); v.value = value; v.font = font(); if(fmt) v.numFmt = fmt;
    v.alignment = {horizontal:"left", vertical:"middle", wrapText:key==="ctx"};
    rr.height = key==="ctx" ? Math.max(20, Math.ceil(String(value).length/150)*16 + 6) : 20;
    if(key) pos[key] = row;
    row++;
  }

  /* títulos das colunas */
  const titulos = [r.porSeg ? "Praça / segmento" : "Praça", "Formato", "CPM tabela", "Desconto", "CPM negociado", "Impressões", "Inventário no período", "Ocupação", "Invest. bruto", "Invest. líquido"];
  const hr = ws.getRow(row); hr.height = 22;
  titulos.forEach((t, i) => { const c = hr.getCell(i+1); c.value = t.toUpperCase(); c.font = font({size:9, bold:true, color:{argb:C.muted}}); c.alignment = {horizontal: i<2 ? "left" : (i===3 ? "center" : "right"), vertical:"middle"}; c.border = linha; });
  row++;

  /* linhas da proposta */
  const D = `$B$${pos.desc}`, DED = `$B$${pos.ded}`;
  const primeira = row;
  r.grupos.forEach((g, gi) => {
    const ls = r.linhas.filter(l=>l.gi===gi), ini = row;
    ls.forEach((l, i) => {
      const rr = ws.getRow(row); rr.height = 20; const n = row;
      const rot = nomePraca(g.p) + (g.s!=="*" ? ` — ${segNome(pub, g.s)}` : "");
      const set = (c, v, fmt, extra={}) => { const cell = rr.getCell(c); cell.value = v; if(fmt) cell.numFmt = fmt; cell.font = font(extra.font||{}); cell.alignment = {vertical:"middle", horizontal: extra.h || (c<=2 ? "left" : "right"), wrapText: c===1}; cell.border = linha; };
      set(1, i===0 ? rot : "", null, {font:{bold:true}});
      set(2, l.fmt.nome);
      set(3, +l.cpmTab.toFixed(2), MONEY);
      set(4, {formula:D, result:r.desc}, '0.0%;-0.0%;"—"', {h:"center", font:{color:{argb:C.muted}}});
      set(5, {formula:`C${n}*(1-D${n})`, result:+l.cpm.toFixed(2)}, MONEY, {font:{bold:true}});
      set(6, {formula:`IF(E${n}>0,J${n}/E${n}*1000,0)`, result:Math.round(l.imps)}, INT);
      set(7, l.invPer==null ? "—" : Math.round(l.invPer), INT);
      set(8, l.invPer==null ? "—" : {formula:`IF(G${n}>0,F${n}/G${n},0)`, result:l.ocup||0}, PC);
      set(9, +l.bruto.toFixed(2), MONEY);
      set(10, {formula:`I${n}*(1-${DED})`, result:+l.liq.toFixed(2)}, MONEY);
      row++;
    });
    if(r.grupos.length > 1){
      const rr = ws.getRow(row), a = ini, b = row-1, n = row;
      const nomeSub = [g.p!=="BR" ? g.p : "", g.s!=="*" ? "segmento " + g.s : ""].filter(Boolean).join(" · ");
      const sub = (c, v, fmt) => { const cell = rr.getCell(c); cell.value = v; if(fmt) cell.numFmt = fmt; cell.font = font({bold:true, size:10}); cell.alignment = {vertical:"middle", horizontal: c===1 ? "left" : "right"}; };
      for(let c=1; c<=N; c++){ rr.getCell(c).fill = fill(C.sunken); rr.getCell(c).border = linha; }
      sub(1, `Subtotal ${nomeSub}`);
      sub(6, {formula:`SUBTOTAL(9,F${a}:F${b})`}, INT);
      sub(7, {formula:`SUBTOTAL(9,G${a}:G${b})`}, INT);
      sub(8, {formula:`IF(G${n}>0,F${n}/G${n},0)`}, PC);
      sub(9, {formula:`SUBTOTAL(9,I${a}:I${b})`}, MONEY);
      sub(10, {formula:`SUBTOTAL(9,J${a}:J${b})`}, MONEY);
      rr.height = 20; row++;
    }
  });
  const ultima = row-1;

  /* total */
  const tr = ws.getRow(row), tn = row; tr.height = 24;
  for(let c=1; c<=N; c++){ const cell = tr.getCell(c); cell.fill = fill(C.soft); cell.font = font({bold:true}); cell.alignment = {vertical:"middle", horizontal: c===1 ? "left" : "right"}; }
  tr.getCell(1).value = "Total";
  tr.getCell(6).value = {formula:`SUBTOTAL(9,F${primeira}:F${ultima})`, result:Math.round(r.tot.imps)}; tr.getCell(6).numFmt = INT;
  if(r.tot.inv!=null){ tr.getCell(7).value = {formula:`SUBTOTAL(9,G${primeira}:G${ultima})`, result:Math.round(r.tot.inv)}; tr.getCell(7).numFmt = INT;
    tr.getCell(8).value = {formula:`IF(G${tn}>0,F${tn}/G${tn},0)`, result:r.tot.ocup||0}; tr.getCell(8).numFmt = PC; }
  tr.getCell(9).value = {formula:`SUBTOTAL(9,I${primeira}:I${ultima})`, result:+r.tot.bruto.toFixed(2)}; tr.getCell(9).numFmt = MONEY;
  tr.getCell(10).value = {formula:`SUBTOTAL(9,J${primeira}:J${ultima})`, result:+r.tot.liq.toFixed(2)}; tr.getCell(10).numFmt = MONEY;
  row++;
  if(r.naoAlocado){
    const wr = ws.getRow(row); wr.height = 20;
    for(let c=1; c<=N; c++){ const cell = wr.getCell(c); cell.fill = fill(C.warnSoft); cell.font = font({color:{argb:C.warn}}); }
    wr.getCell(1).value = "Verba sem entrega (não cabe no inventário)";
    wr.getCell(9).value = +r.naoAlocado.toFixed(2); wr.getCell(9).numFmt = MONEY; wr.getCell(9).alignment = {horizontal:"right"};
    row++;
  }

  /* barra verde na ocupação, como na tela */
  try{ ws.addConditionalFormatting({ref:`H${primeira}:H${ultima}`, rules:[{type:"dataBar", minLength:0, maxLength:100, gradient:false, showValue:true, cfvo:[{type:"num", value:0},{type:"num", value:r.fatorTeto}], color:{argb:C.green}}]}); }catch(e){}

  /* notas */
  row++;
  const notas = [
    `CPM negociado = CPM tabela × (1 − desconto). Impressões = investimento líquido ÷ CPM negociado × 1.000. Cada linha entrega no máximo ${pct(r.fatorTeto)} do inventário do período.`,
    r.inv.fonte==="gam" ? `Inventário: Google Ad Manager, importado em ${fmtData(pub.inventario.importadoEm)} (média de ${pub.inventario.dias} dias).` : r.inv.fonte==="estimado" ? "Inventário estimado por pageviews." : "Sem inventário importado."
  ];
  for(const t of notas){ ws.mergeCells(`A${row}:${LAST}${row}`); const c = ws.getCell(`A${row}`); c.value = t; c.font = font({size:9, color:{argb:C.muted}}); c.alignment = {wrapText:true, vertical:"top"}; ws.getRow(row).height = 26; row++; }

  const buf = await wb.xlsx.writeBuffer();
  const nome = `Proposta ${pub.nome} - ${f.cliente||"anunciante"}`.replace(/[\\/:*?"<>|]+/g," ").slice(0,120) + ".xlsx";
  try{ await S.dl.save({filename:nome, data:buf}); }
  catch(e){ if(e && e.code==="host_unavailable"){ S.dl = null; render(); } }
}
/* ---------- apresentação (PowerPoint) com a marca do publisher ---------- */
function pptCores(cor){
  const ok = /^#[0-9a-f]{6}$/i.test(String(cor||"")) ? String(cor).toUpperCase() : "#8BC53F";
  const rgb = h => [1,3,5].map(i=>parseInt(h.slice(i,i+2),16));
  const hx = a => a.map(v=>Math.round(Math.max(0,Math.min(255,v))).toString(16).padStart(2,"0")).join("").toUpperCase();
  const mix = (a,b,t) => { const x = rgb(a), y = rgb(b); return "#" + hx(x.map((v,i)=>v+(y[i]-v)*t)); };
  const lum = h => { const c = rgb(h).map(v=>{ v/=255; return v<=0.03928 ? v/12.92 : Math.pow((v+0.055)/1.055,2.4); }); return 0.2126*c[0]+0.7152*c[1]+0.0722*c[2]; };
  const con = (a,b) => { const x = lum(a), y = lum(b); return (Math.max(x,y)+0.05)/(Math.min(x,y)+0.05); };
  let forte = ok; for(let t=0.05; t<=0.95 && con(forte,"#FFFFFF")<4.6; t+=0.05) forte = mix(ok,"#000000",t);
  const s = h => h.slice(1);
  return {accent:s(ok), strong:s(forte), soft:s(mix(ok,"#FFFFFF",0.86)), sunken:s(mix(ok,"#FFFFFF",0.94)), line:s(mix(mix(ok,"#FFFFFF",0.8),"#DADADA",0.55)),
    ink: con(ok,"#1A1A1A") >= con(ok,"#FFFFFF") ? "1A1A1A" : "FFFFFF", fg:"1A1A1A", muted:s(mix(forte,"#5A5A5A",0.72))};
}
async function pptLogo(pub){
  const m = pub.marca || {};
  let url = logoUrl(m);
  if(!url){ const el = document.querySelector(".brand img"); if(el && !el.hidden && el.getAttribute("src")) url = el.getAttribute("src"); }
  if(!url) return null;
  try{
    const resp = await fetch(url);
    if(resp.ok){
      const u8 = new Uint8Array(await resp.arrayBuffer());
      const tipo = u8[0]===0x89 && u8[1]===0x50 ? "png" : u8[0]===0xFF && u8[1]===0xD8 ? "jpeg" : null;
      if(tipo){
        let bin = ""; for(let i=0; i<u8.length; i+=0x8000) bin += String.fromCharCode.apply(null, u8.subarray(i, i+0x8000));
        const data = `data:image/${tipo};base64,` + btoa(bin);
        const ar = await new Promise(r2 => { const im = new Image(); im.onload = () => r2((im.naturalWidth||1)/(im.naturalHeight||1)); im.onerror = () => r2(3.6); im.src = data; });
        return {data, ar};
      }
    }
  }catch(e){}
  return new Promise(res => {
    const im = new Image();
    im.onload = () => { try{ const k = Math.min(1, 1200/(im.naturalWidth||1)); const c = document.createElement("canvas"); c.width = Math.max(1, Math.round(im.naturalWidth*k)); c.height = Math.max(1, Math.round(im.naturalHeight*k)); c.getContext("2d").drawImage(im, 0, 0, c.width, c.height); res({data:c.toDataURL("image/png"), ar:c.width/c.height}); }catch(e){ res(null); } };
    im.onerror = () => res(null); im.src = url;
  });
}
async function exportarPPT(pub, f, r){
  if(!S.dl || !window.PptxGenJS) return;
  const C = pptCores((pub.marca||{}).cor), logo = await pptLogo(pub);
  const FONT = "Arial", W = 13.333, H = 7.5;
  const pptx = new PptxGenJS(); pptx.layout = "LAYOUT_WIDE";
  pptx.title = `Proposta ${pub.nome} - ${f.cliente||"anunciante"}`; pptx.company = pub.nome; pptx.subject = "Proposta de mídia";
  const segsTxt = r.porSeg ? r.segs.map(id=>segNome(pub,id)).join("; ") : "Todos os segmentos";
  const pracasTxt = r.geo ? r.pracas.map(p=>UF[p] ? UF[p].nome : p).join(", ") : "Nacional";
  const periodo = `${fmtData(f.inicio)} a ${fmtData(f.fim)} · ${f.dias} dias`;
  const cpmMedio = r.tot.imps ? r.tot.liq/r.tot.imps*1000 : 0;
  let pagina = 0;
  const logoBox = (s, x, y, maxW, maxH) => { if(!logo) return false; let w = maxW, h = w/logo.ar; if(h > maxH){ h = maxH; w = h*logo.ar; } s.addImage({data:logo.data, x:x, y:y + (maxH-h)/2, w, h}); return true; };
  const base = (titulo, sub) => {
    const s = pptx.addSlide(); pagina++;
    s.background = {color:"FFFFFF"};
    s.addShape(pptx.ShapeType.rect, {x:0, y:0, w:0.18, h:H, fill:{color:C.accent}, line:{color:C.accent}});
    s.addText(titulo, {x:0.6, y:0.38, w:9.6, h:0.7, fontFace:FONT, fontSize:28, bold:true, color:C.fg});
    if(sub) s.addText(sub, {x:0.6, y:1.02, w:9.6, h:0.4, fontFace:FONT, fontSize:14, color:C.muted, fit:"shrink"});
    if(!logoBox(s, W-2.6, 0.4, 2.0, 0.65)) s.addText(pub.nome, {x:W-3.4, y:0.4, w:2.8, h:0.65, fontFace:FONT, fontSize:16, bold:true, color:C.strong, align:"right"});
    s.addShape(pptx.ShapeType.line, {x:0.6, y:1.5, w:W-1.2, h:0, line:{color:C.line, width:1}});
    s.addText(`${f.cliente||"Anunciante"} · ${pub.nome}`, {x:0.6, y:H-0.5, w:8, h:0.3, fontFace:FONT, fontSize:10, color:C.muted});
    s.addText(String(pagina), {x:W-1.2, y:H-0.5, w:0.6, h:0.3, fontFace:FONT, fontSize:10, color:C.muted, align:"right"});
    return s;
  };
  const cartao = (s, x, y, w, h, rotulo, valor, nota) => {
    s.addShape(pptx.ShapeType.roundRect, {x, y, w, h, rectRadius:0.08, fill:{color:C.sunken}, line:{color:C.line, width:0.75}});
    s.addText(rotulo.toUpperCase(), {x:x+0.25, y:y+0.18, w:w-0.5, h:0.3, fontFace:FONT, fontSize:10, bold:true, color:C.muted, charSpacing:1});
    s.addText(valor, {x:x+0.25, y:y+0.5, w:w-0.5, h:0.7, fontFace:FONT, fontSize:26, bold:true, color:C.strong, fit:"shrink"});
    if(nota) s.addText(nota, {x:x+0.25, y:y+1.18, w:w-0.5, h:0.4, fontFace:FONT, fontSize:11, color:C.muted});
  };

  /* 1. capa */
  { const s = pptx.addSlide(); pagina++;
    s.background = {color:"FFFFFF"};
    s.addShape(pptx.ShapeType.rect, {x:0, y:0, w:7.2, h:H, fill:{color:C.accent}, line:{color:C.accent}});
    s.addText("PROPOSTA DE MÍDIA", {x:0.7, y:1.6, w:6, h:0.4, fontFace:FONT, fontSize:14, bold:true, color:C.ink, charSpacing:2});
    s.addText(f.cliente || "Anunciante", {x:0.7, y:2.1, w:6.1, h:2.4, fontFace:FONT, fontSize:40, bold:true, color:C.ink, valign:"top", fit:"shrink"});
    if(!logoBox(s, 8.2, 2.6, 4.2, 1.6)) s.addText(pub.nome, {x:7.8, y:2.8, w:5, h:1.2, fontFace:FONT, fontSize:36, bold:true, color:C.strong, align:"center"});
    s.addText(`Apresentado por ${pub.nome}`, {x:7.8, y:4.5, w:5, h:0.4, fontFace:FONT, fontSize:14, color:C.muted, align:"center"});
    s.addText(new Date().toLocaleDateString("pt-BR",{day:"2-digit", month:"long", year:"numeric"}), {x:7.8, y:H-0.9, w:5, h:0.35, fontFace:FONT, fontSize:11, color:C.muted, align:"center"});
  }

  /* 2. desafio */
  { const s = base("O desafio da campanha", "O que a campanha precisa comunicar e para quem");
    s.addShape(pptx.ShapeType.rect, {x:0.6, y:1.95, w:0.08, h:2.2, fill:{color:C.accent}, line:{color:C.accent}});
    s.addText(f.contexto || "Contexto a definir com o anunciante.", {x:0.95, y:1.9, w:11.5, h:2.3, fontFace:FONT, fontSize:22, color:C.fg, valign:"middle", fit:"shrink"});
    const itens = [["Período", periodo], ["Investimento bruto", BRL.format(f.verba)], ["Praças", pracasTxt], ["Público", segsTxt]];
    itens.forEach(([k,v], i) => { const x = 0.6 + i*3.06;
      s.addShape(pptx.ShapeType.roundRect, {x, y:4.7, w:2.9, h:1.7, rectRadius:0.08, fill:{color:C.sunken}, line:{color:C.line, width:0.75}});
      s.addText(k.toUpperCase(), {x:x+0.2, y:4.85, w:2.5, h:0.3, fontFace:FONT, fontSize:10, bold:true, color:C.muted, charSpacing:1});
      s.addText(v, {x:x+0.2, y:5.2, w:2.5, h:1.05, fontFace:FONT, fontSize:14, bold:true, color:C.fg, valign:"top", fit:"shrink"}); });
  }

  /* 3. por que o publisher */
  { const inv = pub.inventario && pub.inventario.fonte==="gam" ? pub.inventario.porFormatoUF || {} : null;
    const s = base(`Por que ${pub.nome}`, inv ? `Audiência medida no Google Ad Manager · média de ${pub.inventario.dias} dias` : "Audiência estimada");
    let dia = 0; const porUF = {};
    if(inv) for(const fid in inv) for(const uf in inv[fid]){ const v = +inv[fid][uf]||0; dia += v; if(uf!=="ND") porUF[uf] = (porUF[uf]||0) + v; }
    const alvo = r.tot.inv!=null ? r.tot.inv/f.dias : null;
    cartao(s, 0.6, 1.9, 3.9, 1.7, "Impressões por mês", dia ? NUM.format(dia*30) : "—", dia ? `${NUM.format(dia)} por dia` : "importe o relatório do GAM");
    cartao(s, 0.6, 3.8, 3.9, 1.7, "No recorte da campanha", alvo!=null ? NUM.format(alvo*30) : "—", alvo!=null && dia ? `${pct(alvo/dia)} da audiência, por mês` : (r.geo||r.porSeg ? "praças e público escolhidos" : "toda a audiência"));
    const top = Object.entries(porUF).sort((a,b)=>b[1]-a[1]).slice(0,8);
    if(top.length){
      s.addText("Onde a audiência está (impressões por dia)", {x:5.0, y:1.75, w:7.7, h:0.4, fontFace:FONT, fontSize:13, bold:true, color:C.fg});
      s.addChart(pptx.ChartType.bar, [{name:"Impressões por dia", labels:top.map(([uf])=>UF[uf] ? UF[uf].nome : uf).reverse(), values:top.map(([,v])=>Math.round(v)).reverse()}],
        {x:5.0, y:2.15, w:7.7, h:4.6, barDir:"bar", chartColors:[C.accent], showValue:true, dataLabelFormatCode:"#,##0", dataLabelFontSize:10, dataLabelColor:C.fg,
         catAxisLabelFontSize:11, catAxisLabelColor:C.fg, catAxisLabelFontFace:FONT, valAxisHidden:true, valGridLine:{style:"none"}, catAxisLineShow:false, showLegend:false, barGapWidthPct:40});
    } else s.addText("O detalhamento por estado aparece depois da importação do relatório do GAM.", {x:5.0, y:3.5, w:7.7, h:0.8, fontFace:FONT, fontSize:14, color:C.muted});
  }

  /* agregações por formato e por grupo */
  const porFmt = {}; for(const l of r.linhas){ const k = l.fmt.id; porFmt[k] ??= {nome:l.fmt.nome, imps:0, bruto:0, liq:0, inv:0, cpmTab:l.cpmTab, cpm:l.cpm}; porFmt[k].imps += l.imps; porFmt[k].bruto += l.bruto; porFmt[k].liq += l.liq; porFmt[k].inv += l.invPer||0; }
  const fmtLista = Object.values(porFmt).filter(x=>x.bruto>0.005).sort((a,b)=>b.bruto-a.bruto);

  /* 4. estratégia */
  { const s = base("Estratégia de mídia", `${fmtLista.length} formato${fmtLista.length>1?"s":""} · ${pracasTxt} · ${segsTxt}`);
    const n = Math.max(1, fmtLista.length), cols = Math.min(n, 3), w = (W-1.2-(cols-1)*0.25)/cols, hh = n>3 ? 2.15 : 3.0;
    fmtLista.slice(0,6).forEach((x, i) => { const col = i%cols, row = Math.floor(i/cols), X = 0.6 + col*(w+0.25), Y = 1.9 + row*(hh+0.25);
      s.addShape(pptx.ShapeType.roundRect, {x:X, y:Y, w, h:hh, rectRadius:0.08, fill:{color:"FFFFFF"}, line:{color:C.line, width:1}});
      s.addShape(pptx.ShapeType.rect, {x:X, y:Y, w, h:0.1, fill:{color:C.accent}, line:{color:C.accent}});
      s.addText(x.nome, {x:X+0.25, y:Y+0.25, w:w-0.5, h:0.5, fontFace:FONT, fontSize:18, bold:true, color:C.fg, fit:"shrink"});
      const share = x.bruto/(r.tot.bruto||1);
      s.addText(share>0 && share<0.01 ? "<1%" : pct(share), {x:X+0.25, y:Y+0.72, w:1.9, h:0.6, fontFace:FONT, fontSize:28, bold:true, color:C.strong, valign:"middle"});
      s.addText("da verba", {x:X+2.1, y:Y+0.72, w:w-2.35, h:0.6, fontFace:FONT, fontSize:12, color:C.muted, valign:"middle"});
      s.addText(`${NUM.format(x.imps)} impressões\nCPM ${BRL.format(x.cpm)}${r.desc ? ` (tabela ${BRL.format(x.cpmTab)})` : ""}`, {x:X+0.25, y:Y+1.38, w:w-0.5, h:hh-1.5, fontFace:FONT, fontSize:12, color:C.fg, valign:"top"});
    });
    if(fmtLista.length > 6) s.addText(`+ ${fmtLista.length-6} formatos no plano de mídia`, {x:0.6, y:H-1.0, w:6, h:0.3, fontFace:FONT, fontSize:11, color:C.muted});
  }

  /* 5. entrega projetada */
  { const s = base("Entrega projetada", r.naoAlocado ? `O inventário comporta ${BRL.format(r.tot.bruto)} da verba nestas praças e públicos` : "Impressões que o inventário entrega no período");
    const w = (W-1.2-0.75)/4;
    cartao(s, 0.6, 1.95, w, 1.9, "Impressões", NUM.format(r.tot.imps), `${NUM.format(r.tot.imps/f.dias)} por dia`);
    cartao(s, 0.6+(w+0.25), 1.95, w, 1.9, "Investimento bruto", BRL.format(r.tot.bruto), r.desc ? `desconto de ${pct(r.desc)} no CPM` : "CPM tabela");
    cartao(s, 0.6+2*(w+0.25), 1.95, w, 1.9, "Investimento líquido", BRL.format(r.tot.liq), `após ${pct(r.ded)} de agência e custos`);
    cartao(s, 0.6+3*(w+0.25), 1.95, w, 1.9, "CPM médio", BRL.format(cpmMedio), "sobre o investimento líquido");
    if(fmtLista.length){
      s.addText("Impressões por formato", {x:0.6, y:4.15, w:6, h:0.35, fontFace:FONT, fontSize:13, bold:true, color:C.fg});
      s.addChart(pptx.ChartType.bar, [{name:"Impressões", labels:fmtLista.map(x=>x.nome), values:fmtLista.map(x=>Math.round(x.imps))}],
        {x:0.6, y:4.5, w:W-1.2, h:2.35, barDir:"col", chartColors:[C.accent], showValue:true, dataLabelFormatCode:"#,##0", dataLabelFontSize:10, dataLabelColor:C.fg, dataLabelPosition:"outEnd",
         catAxisLabelFontSize:11, catAxisLabelColor:C.fg, catAxisLabelFontFace:FONT, valAxisHidden:true, valGridLine:{style:"none"}, showLegend:false, barGapWidthPct:60});
    }
  }

  /* 6. plano de mídia */
  { const s = base("Plano de mídia", `${periodo} · ${pracasTxt}`);
    const hd = ["Formato","CPM tabela","Desconto","CPM negociado","Impressões","Invest. bruto","Invest. líquido"];
    const th = t => ({text:t, options:{bold:true, fontSize:10, color:C.muted, fill:{color:C.sunken}, align: t==="Formato" ? "left" : "right"}});
    const td = (t, o={}) => ({text:t, options:{fontSize:11, color:C.fg, align:"right", ...o}});
    const rows = [hd.map(th)];
    for(const x of fmtLista) rows.push([td(x.nome,{align:"left", bold:true}), td(BRL.format(x.cpmTab)), td(r.desc ? pct(r.desc) : "—"), td(BRL.format(x.cpm),{bold:true}), td(NUM.format(x.imps)), td(BRL.format(x.bruto)), td(BRL.format(x.liq))]);
    const tot = {fill:{color:C.soft}, bold:true};
    rows.push([td("Total",{align:"left", ...tot}), td("",tot), td("",tot), td(BRL.format(cpmMedio),tot), td(NUM.format(r.tot.imps),tot), td(BRL.format(r.tot.bruto),tot), td(BRL.format(r.tot.liq),tot)]);
    if(r.naoAlocado) rows.push([td("Verba sem entrega (não cabe no inventário)",{align:"left", color:"9A6414", colspan:5}), td(BRL.format(r.naoAlocado),{color:"9A6414"}), td("")]);
    s.addTable(rows, {x:0.6, y:1.9, w:W-1.2, colW:[3.4,1.35,1.15,1.6,1.6,1.5,1.53], fontFace:FONT, border:{type:"solid", pt:0.5, color:C.line}, rowH:0.42, autoPage:false});
    if(r.grupos.length > 1) s.addText(`Detalhamento por ${r.geo && r.porSeg ? "praça e segmento" : r.geo ? "praça" : "segmento"} na proposta em Excel.`, {x:0.6, y:H-1.0, w:9, h:0.3, fontFace:FONT, fontSize:11, color:C.muted});
  }

  /* 7. distribuição por praça/segmento (quando houver mais de um grupo) */
  if(r.grupos.length > 1){
    const g = {}; r.linhas.forEach(l => { const k = (l.praca==="BR" ? "" : (UF[l.praca] ? UF[l.praca].nome : l.praca)) + (l.seg!=="*" ? (l.praca==="BR" ? "" : " · ") + segNome(pub,l.seg).replace(/^\d+ · /,"") : ""); g[k] = (g[k]||0) + l.imps; });
    const lista = Object.entries(g).sort((a,b)=>b[1]-a[1]).slice(0,12);
    const s = base("Distribuição da entrega", `Impressões por ${r.geo && r.porSeg ? "praça e segmento" : r.geo ? "praça" : "segmento"}`);
    s.addChart(pptx.ChartType.bar, [{name:"Impressões", labels:lista.map(x=>x[0]).reverse(), values:lista.map(x=>Math.round(x[1])).reverse()}],
      {x:0.6, y:1.8, w:W-1.2, h:5.0, barDir:"bar", chartColors:[C.accent], showValue:true, dataLabelFormatCode:"#,##0", dataLabelFontSize:10, dataLabelColor:C.fg,
       catAxisLabelFontSize:11, catAxisLabelColor:C.fg, catAxisLabelFontFace:FONT, valAxisHidden:true, valGridLine:{style:"none"}, showLegend:false, barGapWidthPct:40});
  }

  /* 8. condições e próximos passos */
  { const s = base("Condições e próximos passos");
    const cond = [
      `Investimento bruto de ${BRL.format(f.verba)}${r.desc ? `, com desconto de ${pct(r.desc)} sobre o CPM tabela` : ", em CPM tabela"}.`,
      `Investimento líquido de ${BRL.format(r.tot.liq)}, após ${pct(r.ded)} de agência e demais custos.`,
      `Entrega estimada de ${NUM.format(r.tot.imps)} impressões, limitada a ${pct(r.fatorTeto)} do inventário de cada formato no período.`,
      pub.inventario && pub.inventario.fonte==="gam" ? `Inventário do Google Ad Manager atualizado em ${fmtData(pub.inventario.importadoEm)}.` : "Inventário estimado; os volumes serão confirmados com dados do Google Ad Manager.",
      "Proposta válida por 15 dias, sujeita à disponibilidade de inventário."
    ];
    s.addText("Condições", {x:0.6, y:1.8, w:6, h:0.4, fontFace:FONT, fontSize:16, bold:true, color:C.strong});
    s.addText(cond.map(t => ({text:t, options:{bullet:{indent:16}, breakLine:true}})), {x:0.6, y:2.25, w:6.6, h:4.4, fontFace:FONT, fontSize:13, color:C.fg, valign:"top", paraSpaceAfter:8});
    const passos = ["Aprovação da proposta", "Envio das peças nos formatos do plano", "Configuração e testes da campanha", "Início da veiculação e relatórios de acompanhamento"];
    s.addText("Próximos passos", {x:7.6, y:1.8, w:5, h:0.4, fontFace:FONT, fontSize:16, bold:true, color:C.strong});
    passos.forEach((t, i) => { const y = 2.3 + i*1.05;
      s.addShape(pptx.ShapeType.ellipse, {x:7.6, y, w:0.55, h:0.55, fill:{color:C.accent}, line:{color:C.accent}});
      s.addText(String(i+1), {x:7.6, y, w:0.55, h:0.55, fontFace:FONT, fontSize:14, bold:true, color:C.ink, align:"center", valign:"middle"});
      s.addText(t, {x:8.35, y, w:4.4, h:0.55, fontFace:FONT, fontSize:14, color:C.fg, valign:"middle"}); });
  }

  /* 9. encerramento */
  { const s = pptx.addSlide(); pagina++;
    s.background = {color:C.accent};
    s.addShape(pptx.ShapeType.roundRect, {x:W/2-3, y:1.7, w:6, h:2.4, rectRadius:0.12, fill:{color:"FFFFFF"}, line:{color:"FFFFFF"}});
    if(!logoBox(s, W/2-2.3, 2.1, 4.6, 1.6)) s.addText(pub.nome, {x:W/2-2.8, y:2.3, w:5.6, h:1.2, fontFace:FONT, fontSize:36, bold:true, color:C.strong, align:"center"});
    s.addText("Obrigado", {x:0, y:4.5, w:W, h:0.8, fontFace:FONT, fontSize:34, bold:true, color:C.ink, align:"center"});
    s.addText(pub.contato ? String(pub.contato) : `${pub.nome}${pub.dominio ? " · " + pub.dominio : ""}`, {x:0, y:5.3, w:W, h:0.5, fontFace:FONT, fontSize:16, color:C.ink, align:"center"});
  }

  const buf = await pptx.write({outputType:"arraybuffer"});
  const nome = `Apresentação ${pub.nome} - ${f.cliente||"anunciante"}`.replace(/[\\/:*?"<>|]+/g," ").slice(0,120) + ".pptx";
  try{ await S.dl.save({filename:nome, data:buf}); }
  catch(e){ if(e && e.code==="host_unavailable"){ S.dl = null; render(); } }
}

function exportarPlano(p, tipo){
  const pub = S.publishers[p.publisherId]; if(!pub) return;
  const f = {cliente:p.cliente, inicio:p.inicio, fim:p.fim, dias:p.dias, verba:p.verba, contexto:p.contexto, modoPraca:p.modoPraca, ufs:p.ufs||[], mix:p.mix||{}, cpm:p.cpm||{}, desconto:+p.desconto||0, modoSeg:p.modoSeg||"todos", segs:p.segs||[], fmtsSel:Array.isArray(p.formatosSel)?p.formatosSel:null};
  const r = calcular(pub, f); if(!r.erro) (tipo==="ppt" ? exportarPPT : exportar)(pub, f, r);
}

/* ---------- importação do relatório do GAM ---------- */
function parseCSV(text){
  text = text.replace(/^\uFEFF/,"");
  const amostra = text.split(/\r?\n/).slice(0,20).join("\n");
  const cnt = d => amostra.split(d).length - 1;
  const delim = [",",";","\t"].sort((a,b)=>cnt(b)-cnt(a))[0];
  const rows = []; let row = [], cell = "", q = false;
  for(let i=0;i<text.length;i++){
    const c = text[i];
    if(q){ if(c==='"'){ if(text[i+1]==='"'){ cell+='"'; i++; } else q=false; } else cell+=c; }
    else if(c==='"') q = true;
    else if(c===delim){ row.push(cell); cell=""; }
    else if(c==="\n"){ row.push(cell); rows.push(row); row=[]; cell=""; }
    else if(c!=="\r") cell += c;
  }
  if(cell!=="" || row.length){ row.push(cell); rows.push(row); }
  return rows;
}
const PADROES = {
  unidade:[/unidade de anuncio(?!.*\bid\b)/, /^ad unit$/, /ad unit(?!.*\bid\b)/, /bloco de anuncio/],
  regiao:[/^regiao$/, /^region$/, /^estado$/, /^state$/, /regiao(?!.*\bid\b)/, /region(?!.*\bid\b)/],
  pais:[/^pais$/, /^country$/, /pais(?!.*\bid\b)/, /country(?!.*\bid\b)/],
  data:[/^data$/, /^date$/, /^dia$/],
  imps:[/total de impress/, /total impressions/, /impressoes totais/, /impress/],
  segmento:[/^segmento$/, /segment/, /chave.?valor/, /key.?value/, /valores? de chave/, /segmentacao personalizada/, /custom targeting/, /dimensao personalizada/, /custom dimension/]
};
function detectarColunas(h){
  const n = h.map(norm), pick = ps => { for(const p of ps){ const i = n.findIndex(x=>p.test(x)); if(i>=0) return i; } return -1; };
  return {unidade:pick(PADROES.unidade), regiao:pick(PADROES.regiao), pais:pick(PADROES.pais), data:pick(PADROES.data), imps:pick(PADROES.imps), segmento:pick(PADROES.segmento)};
}
const toNum = s => { let t = String(s??"").trim(); if(!t) return 0; t = t.replace(/[.,]\d{1,2}$/,""); const d = t.replace(/[^\d]/g,""); return d ? parseInt(d,10) : 0; };
function ufDe(s){
  let n = norm(s).replace(/^br-/,"").replace(/,?\s*(brazil|brasil)$/,"").replace(/^(state of|estado de|estado do|estado da|estado)\s+/,"").trim();
  if(/^[a-z]{2}$/.test(n) && UF[n.toUpperCase()]) return n.toUpperCase();
  return UF_POR_NOME[n] || "ND";
}
function segDe(v, pub){
  let s = String(v ?? "").trim(); if(!s || /^\(?(not set|nao definido|n\/a|-)\)?$/i.test(norm(s))) return {seg:"ND"};
  let key = null; const m = s.match(/^([^=]+)=(.*)$/);
  if(m){ key = norm(m[1]); s = m[2].trim(); }
  const chave = norm(pub.chaveSegmento||"");
  if(key && chave && key!==chave) return {outra:true};
  const segs = pub.segmentos||[];
  const num = (s.match(/^\d+/)||[])[0];
  if(num && segs.some(x=>String(x.id)===num)) return {seg:num};
  const n = norm(s), porNome = segs.find(x => norm(x.nome) && (n===norm(x.nome) || n.includes(norm(x.nome))));
  if(porNome) return {seg:String(porNome.id)};
  if(key && !chave && !/segment/.test(key)) return {outra:true};
  return {seg:"ND"};
}
const blocoDe = u => String(u).split(/[»>\/|]/).pop().trim() || String(u);
const nomeBloco = k => { const t = String(k).trim(); return t.charAt(0).toUpperCase() + t.slice(1); };
function casarFormato(u, pub){
  const un = norm(u);
  if((pub.ignorar||[]).some(k=>k && un.includes(norm(k)))) return "__ignorar";
  let best = null, len = 0;
  for(const f of (pub.formatos||[]).filter(x=>x.ativo!==false)) for(const k of (f.chaves||[])){ const nk = norm(k); if(nk && un.includes(nk) && nk.length > len){ best = f.id; len = nk.length; } }
  return best;
}
async function lerArquivo(file){
  const ext = file.name.split(".").pop().toLowerCase();
  if(ext==="xlsx" || ext==="xls"){
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, {type:"array"});
    return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header:1, raw:false, defval:""});
  }
  return parseCSV(await file.text());
}
async function iniciarImportacao(file){
  S.imp = {erro:"", arquivo:file.name};
  try{
    const rows = await lerArquivo(file);
    let hi = -1, cols = null;
    for(let i=0;i<Math.min(rows.length,30);i++){ const c = detectarColunas(rows[i]); if(c.unidade>=0 && c.imps>=0){ hi = i; cols = c; break; } }
    if(hi<0){ S.imp = {arquivo:file.name, erro:"Não encontrei as colunas de unidade de anúncio e de impressões. Confira se o relatório tem a dimensão Unidade de anúncio e a métrica Total de impressões."}; renderPublisher(); return; }
    S.imp = {arquivo:file.name, header:rows[hi], rows:rows.slice(hi+1), cols, atrib:{}, diasManual:30, erro:""};
    agregar();
  } catch(e){ S.imp = {arquivo:file.name, erro:"Não consegui ler o arquivo. Exporte o relatório do GAM em CSV ou Excel e tente de novo."}; }
  renderPublisher();
}
function agregar(){
  const I = S.imp, pub = S.draft, c = I.cols;
  const usaSeg = c.segmento>=0 && (pub.segmentos||[]).length>0;
  const acc = {}, cubo = {}, segTot = {}, semFmt = {}, datas = new Set();
  let total=0, ignorado=0, foraBR=0, semReg=0, semSeg=0, outras=0, linhas=0;
  for(const row of I.rows){
    const u = String(row[c.unidade]??"").trim();
    if(!u || /^total/i.test(u)) continue;
    const imps = toNum(row[c.imps]); if(!imps) continue;
    if(c.pais>=0){ const p = norm(row[c.pais]); if(p && !/^(brazil|brasil|br)$/.test(p)){ foraBR += imps; continue; } }
    let seg = null;
    if(usaSeg){ const sr = segDe(row[c.segmento], pub); if(sr.outra){ outras += imps; continue; } seg = sr.seg; }
    if(c.data>=0 && String(row[c.data]).trim()) datas.add(String(row[c.data]).trim());
    linhas++; total += imps;
    let fid = I.atrib[u] ?? casarFormato(u, pub);
    if(fid==="__novo") fid = "blk-" + slug(blocoDe(u));
    if(fid==null){ semFmt[u] = (semFmt[u]||0) + imps; continue; }
    if(fid==="__ignorar"){ ignorado += imps; continue; }
    const uf = c.regiao>=0 ? ufDe(row[c.regiao]) : "ND";
    if(uf==="ND") semReg += imps;
    (acc[fid] ??= {}); acc[fid][uf] = (acc[fid][uf]||0) + imps;
    if(usaSeg){ ((cubo[fid] ??= {})[uf] ??= {}); cubo[fid][uf][seg] = (cubo[fid][uf][seg]||0) + imps; segTot[seg] = (segTot[seg]||0) + imps; if(seg==="ND") semSeg += imps; }
  }
  const usado = total - ignorado - somaObj(semFmt);
  Object.assign(I, {acc, cubo, segTot, usaSeg, dias: datas.size || I.diasManual, temData: datas.size>0, total, ignorado, foraBR, semReg, semSeg, outras, usado, linhas});
}
function mostrarImportacao(){
  const I = S.imp; if(!I) return "";
  if(I.erro) return `<div class="banner" style="border-color:var(--bad);background:var(--bad-soft)"><span>${esc(I.erro)}</span><span><button class="btn sm" data-imp="cancelar">Fechar</button></span></div>`;
  const opts = (sel) => `<option value="">— escolher —</option>` + (S.draft.formatos||[]).filter(f=>f.ativo!==false).map(f=>`<option value="${esc(f.id)}"${sel===f.id?" selected":""}>${esc(f.nome)}</option>`).join("") + `<option value="__novo"${sel==="__novo"?" selected":""}>+ Criar formato com este bloco</option><option value="__ignorar"${sel==="__ignorar"?" selected":""}>Ignorar (não vendável)</option>`;
  const colSel = (k, label) => `<label>${label}<select class="in" data-col="${k}"><option value="-1">não usar</option>${I.header.map((h,i)=>`<option value="${i}"${I.cols[k]===i?" selected":""}>${esc(h)}</option>`).join("")}</select></label>`;
  const semFmt = Object.entries(I.semFmt).sort((a,b)=>b[1]-a[1]);
  const segsCfg = S.draft.segmentos||[];
  const segResumo = !segsCfg.length ? "" : !I.usaSeg ? `<p class="note" style="color:var(--warn)">O relatório não tem a coluna de segmento. Para ter o volume por segmento, inclua a dimensão Chave-valor filtrada pela chave do segmento, ou escolha a coluna acima.</p>` :
    `<div><b>Impressões por segmento no período</b>${I.outras?` <span class="note">(${NUM.format(I.outras)} impressões de outras chaves descartadas)</span>`:""}
      <div class="tablebox" style="margin-top:8px"><table><thead><tr><th>Segmento</th><th class="num">Impressões</th><th class="num">Por dia</th><th class="num">Participação</th></tr></thead><tbody>${
        [...segsCfg.map(x=>String(x.id)), "ND"].filter(id => id!=="ND" || I.segTot.ND).map(id => { const v = I.segTot[id]||0, tt = somaObj(I.segTot); return `<tr><td>${esc(id==="ND" ? "Sem segmento" : segNome(S.draft,id))}</td><td class="num">${NUM.format(v)}</td><td class="num">${NUM.format(v/I.dias)}</td><td class="num">${tt?pct(v/tt):"—"}</td></tr>`; }).join("")
      }</tbody></table></div></div>`;
  const porFmt = (S.draft.formatos||[]).filter(f=>f.ativo!==false).map(f => `<span>${esc(f.nome)}: <b>${NUM.format(somaObj(I.acc[f.id]))}</b></span>`).join("");
  return `<div class="panel section" style="border-color:var(--accent)">
    <h3>Revisar importação · ${esc(I.arquivo)}</h3>
    <div class="fields">${colSel("unidade","Unidade de anúncio")}${colSel("regiao","Região (estado)")}${colSel("pais","País")}${colSel("data","Data")}${colSel("imps","Impressões")}${(S.draft.segmentos||[]).length ? colSel("segmento","Segmento (chave-valor)") : ""}</div>
    <div class="kv"><span>Linhas lidas: <b>${NUM.format(I.linhas)}</b></span><span>Impressões no Brasil: <b>${NUM.format(I.total)}</b></span><span>Dias: <b>${I.dias}</b></span>${I.foraBR?`<span>Fora do Brasil (descartadas): <b>${NUM.format(I.foraBR)}</b></span>`:""}<span>Sem estado identificado: <b>${I.total?pct(I.semReg/I.total):"0%"}</b></span></div>
    ${I.temData ? "" : `<label class="hint" style="display:flex;gap:8px;align-items:center">O relatório não tem a coluna Data. Quantos dias ele cobre? <input type="number" min="1" id="imp-dias" value="${I.diasManual}" style="width:80px"></label>`}
    ${segResumo}
    <div class="kv">${porFmt}${I.ignorado?`<span>Ignorado: <b>${NUM.format(I.ignorado)}</b></span>`:""}</div>
    ${semFmt.length ? `<div><b>${semFmt.length} unidade${semFmt.length>1?"s":""} de anúncio sem formato</b> <span class="note">(${pct(somaObj(I.semFmt)/I.total)} das impressões, fora do inventário até você escolher)</span>
      <div class="tablebox" style="margin-top:8px;max-height:320px;overflow:auto"><table><thead><tr><th>Unidade de anúncio</th><th class="num">Impressões</th><th>Formato</th></tr></thead><tbody>${
        semFmt.slice(0,80).map(([u,v]) => `<tr><td style="overflow-wrap:anywhere">${esc(u)}</td><td class="num">${NUM.format(v)}</td><td><select class="in" data-atrib="${esc(u)}">${opts(I.atrib[u])}</select></td></tr>`).join("")
      }</tbody></table></div><p class="note" style="margin-top:6px">A escolha vira regra do publisher: nas próximas importações, unidades com o mesmo nome entram no formato automaticamente.</p></div>` : `<span class="chip ok">Todas as unidades de anúncio têm formato</span>`}
    <div style="display:flex;gap:8px;flex-wrap:wrap"><button class="btn primary" data-imp="confirmar" ${I.usado>0?"":"disabled"}>Usar este inventário</button><button class="btn" data-imp="cancelar">Cancelar</button></div>
  </div>`;
}
async function confirmarImportacao(){
  const I = S.imp, d = S.draft;
  for(const [u, fid] of Object.entries(I.atrib)){
    if(!fid) continue;
    const chave = norm(String(u).split(/[»>\/|]/).pop());
    if(!chave) continue;
    if(fid==="__ignorar"){ d.ignorar = [...new Set([...(d.ignorar||[]), chave])]; }
    else if(fid==="__novo"){ const id = "blk-" + slug(blocoDe(u)); if(!d.formatos.some(x=>x.id===id)) d.formatos.push({id, nome:nomeBloco(blocoDe(u)), cpm:0, mix:0, impPorPv:0, ativo:true, chaves:[chave]}); }
    else { const f = d.formatos.find(x=>x.id===fid); if(f && !(f.chaves||[]).includes(chave)) f.chaves = [...(f.chaves||[]), chave]; }
  }
  const porFormatoUF = {};
  for(const [fid, ufs] of Object.entries(I.acc)){ porFormatoUF[fid] = {}; for(const [uf, v] of Object.entries(ufs)) porFormatoUF[fid][uf] = Math.round(v / I.dias); }
  let porFormatoUFSeg = null;
  if(I.usaSeg){ porFormatoUFSeg = {}; for(const [fid, ufs] of Object.entries(I.cubo)){ porFormatoUFSeg[fid] = {}; for(const [uf, ss] of Object.entries(ufs)){ porFormatoUFSeg[fid][uf] = {}; for(const [sg, v] of Object.entries(ss)) porFormatoUFSeg[fid][uf][sg] = Math.round(v / I.dias * 10)/10; } } }
  d.inventario = {porFormatoUFSeg, semSegmentoPct: I.usaSeg && I.usado ? +(I.semSeg/I.usado).toFixed(4) : null, fonte:"gam", arquivo:I.arquivo, importadoEm:new Date().toISOString(), dias:I.dias, totalPeriodo:I.usado, semRegiaoPct: I.total ? +(I.semReg/I.total).toFixed(4) : 0, porFormatoUF};
  S.imp = null; S.dirty = true;
  await salvarPublisher("Inventário do GAM importado e salvo.");
}

/* ---------- configurar publisher ---------- */
function syncDraft(){
  const pub = curPub();
  if(!S.dirty || S.draftFor !== S.pubId){ S.draft = pub ? clone(pub) : null; S.draftFor = S.pubId; S.dirty = false; S.mergeSel = new Set(); }
  if(S.draft && !S.draft.auto) S.draft.auto = AUTO_PADRAO();
}
function passos(d){
  const fmts = (d.formatos||[]).filter(f=>f.ativo!==false);
  const inv = d.inventario && d.inventario.fonte==="gam";
  return [
    {ok: !!(d.nome && d.dominio), t:"Dados do publisher", s: d.dominio ? esc(d.dominio) : "nome e domínio"},
    {ok: fmts.length>0 && fmts.every(f=>+f.cpm>0) && !!d.cpmValidado, t:"Formatos e CPMs", s: !fmts.length ? "nenhum formato" : d.cpmValidado ? `${fmts.length} formatos, CPMs validados` : "CPMs padrão a validar"},
    {ok: inv, t:"Inventário do GAM", s: inv ? `importado em ${esc(fmtData(d.inventario.importadoEm))}` : (+d.pageviewsMes>0 ? "usando estimativa por pageviews" : "relatório ainda não importado")},
    {ok: !!(d.auto && d.auto.planilhaUrl && (d.auto.modo==="api" ? d.networkCode : d.auto.relatorio)), t:"Atualização automática", s: (d.auto && d.auto.modo==="api" && !d.networkCode) ? "código de rede do GAM" : !(d.auto && d.auto.planilhaUrl) ? "agendamento e planilha a configurar" : `${esc(({diaria:"diária",semanal:"semanal",mensal:"mensal"})[d.auto.frequencia]||"semanal")} · rede ${esc(d.networkCode)}`}
  ];
}
function gerarScript(d){
  const a = d.auto || AUTO_PADRAO(), h = Math.min(23, Math.max(0, parseInt(a.hora,10)||6));
  const gat = a.frequencia==="diaria" ? `ScriptApp.newTrigger('atualizarInventario').timeBased().everyDays(1).atHour(${h}).create();`
    : a.frequencia==="mensal" ? `ScriptApp.newTrigger('atualizarInventario').timeBased().onMonthDay(1).atHour(${h}).create();`
    : `ScriptApp.newTrigger('atualizarInventario').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(${h}).create();`;
  const freq = a.frequencia==="diaria" ? "todo dia" : a.frequencia==="mensal" ? "todo dia 1º do mês" : "toda segunda-feira";
  const modo = a.modo==="api" ? "api" : "email";
  return $(modo==="api" ? "#tpl-gas" : "#tpl-mail").textContent.trim()
    .replaceAll("{{RELATORIO}}", String(a.relatorio||"NOME DO RELATÓRIO NO GAM").replace(/["\\]/g,""))
    .replaceAll("{{PUB}}", d.nome||"").replaceAll("{{REDE}}", String(d.networkCode||"CODIGO_DA_REDE").replace(/[^0-9A-Za-z_-]/g,""))
    .replaceAll("{{ABA}}", String(a.aba||"inventario_planejador").replace(/['\\]/g,"")).replaceAll("{{FREQ_TXT}}", freq).replaceAll("{{HORA}}", String(h)).replaceAll("{{GATILHO}}", gat);
}
const MANIFESTO_EMAIL = JSON.stringify({timeZone:"America/Sao_Paulo", runtimeVersion:"V8", exceptionLogging:"STACKDRIVER"}, null, 2);
const MANIFESTO = JSON.stringify({timeZone:"America/Sao_Paulo", runtimeVersion:"V8", exceptionLogging:"STACKDRIVER", oauthScopes:["https://www.googleapis.com/auth/admanager","https://www.googleapis.com/auth/spreadsheets.currentonly","https://www.googleapis.com/auth/script.external_request","https://www.googleapis.com/auth/script.scriptapp"]}, null, 2);
function secaoAuto(d, ro){
  const a = {...AUTO_PADRAO(), ...(d.auto||{})}, dis = ro ? "disabled" : "";
  const api = a.modo==="api";
  const temOrigem = api ? !!d.networkCode : !!a.relatorio, temPlan = !!a.planilhaUrl;
  const freqTxt = ({diaria:"todo dia", semanal:"toda segunda-feira", mensal:"todo dia 1º"})[a.frequencia] || "toda segunda-feira";
  const passos = api ? `
        <li>No Ad Manager do publisher, em Admin › Configurações globais, ative o <b>acesso à API</b>. A conta Google que vai rodar o script precisa ter acesso de leitura a relatórios na rede ${esc(d.networkCode||"do publisher")}.</li>
        <li>No Google Cloud, crie um projeto e ative a <b>Google Ad Manager API</b>; configure a tela de consentimento OAuth.</li>
        <li>Crie a planilha Google "Inventário ${esc(d.nome)}" e abra Extensões › Apps Script de dentro dela.</li>
        <li>Em Configurações do projeto: marque "Mostrar o arquivo de manifesto appsscript.json" e troque o projeto do Google Cloud pelo do passo 2.</li>
        <li>Cole o manifesto em <b>appsscript.json</b> e o script em <b>Código.gs</b> (arquivos separados).</li>
        <li>Rode <b>atualizarInventario</b> uma vez, autorize e confira a aba "${esc(a.aba)}".</li>
        <li>Rode <b>agendar</b> (${esc(freqTxt)} às ${esc(a.hora)}h). Depois cole o link da planilha acima e salve.</li>` : `
        <li>No Ad Manager, abra o relatório de inventário (Data, Bloco de anúncios, País, Região, Chave-valor e Total de impressões, últimos 30 dias) e salve com um nome fixo, por exemplo "${esc(a.relatorio||"Base "+d.nome)}". Informe esse nome acima.</li>
        <li>No mesmo relatório, use <b>Agendar</b>: frequência ${esc(a.frequencia==="diaria"?"diária":a.frequencia==="mensal"?"mensal":"semanal")}, formato <b>CSV</b>, destinatário = o e-mail da conta Google que vai rodar o script.</li>
        <li>Crie a planilha Google "Inventário ${esc(d.nome)}" e abra Extensões › Apps Script de dentro dela.</li>
        <li>Cole o script em <b>Código.gs</b>, substituindo tudo. No <b>appsscript.json</b>, deixe só o manifesto abaixo (sem lista de escopos).</li>
        <li>Quando o primeiro e-mail do GAM chegar, rode <b>atualizarInventario</b>, autorize o acesso ao Gmail e às Planilhas e confira a aba "${esc(a.aba)}".</li>
        <li>Rode <b>agendar</b>: o script passa a rodar ${esc(freqTxt)} às ${esc(a.hora)}h. Programe o horário para depois da chegada do e-mail do GAM. Depois cole o link da planilha acima e salve.</li>`;
  return `<div class="panel section"><h3>Atualização automática do GAM</h3>
    <p class="note">O inventário se atualiza sozinho: o GAM gera o relatório, um script do Google grava o resultado numa planilha e o planejador lê essa planilha e atualiza a base, com as mesmas regras de formatos e segmentos da importação manual.</p>
    <div class="seg" role="radiogroup" aria-label="Como o relatório sai do GAM" style="max-width:520px">
      <input type="radio" name="automodo" id="am-email" value="email" ${api?"":"checked"} ${dis}><label for="am-email">E-mail agendado do GAM (sem Google Cloud)</label>
      <input type="radio" name="automodo" id="am-api" value="api" ${api?"checked":""} ${dis}><label for="am-api">API do Ad Manager</label>
    </div>
    <div class="flow">
      <div class="st ${temOrigem?"on":""}"><b>1 · Relatório no GAM</b><span>${api ? (d.networkCode ? `API, rede ${esc(d.networkCode)}, últimos 30 dias` : "Informe o código de rede em Dados do publisher") : (a.relatorio ? `Agendado por e-mail: "${esc(a.relatorio)}"` : "Informe o nome do relatório agendado")}</span></div>
      <div class="st ${temPlan?"on":""}"><b>2 · Planilha Google</b><span>${temPlan ? `Aba "${esc(a.aba)}", ${esc(freqTxt)} às ${esc(a.hora)}h` : "Cole o link da planilha abaixo"}</span></div>
      <div class="st"><b>3 · Planejador lê a planilha</b><span>${d.inventario && d.inventario.fonte==="gam" ? `Última atualização: ${esc(fmtData(d.inventario.importadoEm))}` : "Ainda sem atualização"} · importação automática pelo servidor: próxima versão</span></div>
    </div>
    <div class="fields">
      ${api ? "" : `<label>Nome do relatório agendado no GAM<input type="text" id="au-rel" data-ak="relatorio" value="${esc(a.relatorio)}" ${dis} placeholder="ex.: Base MUP"><span class="hint">O script procura e-mails com esse nome no assunto.</span></label>`}
      <label>Frequência<select class="in" id="au-freq" data-ak="frequencia" ${dis}>${[["diaria","Diária"],["semanal","Semanal (segunda-feira)"],["mensal","Mensal (dia 1º)"]].map(([v,t])=>`<option value="${v}"${a.frequencia===v?" selected":""}>${t}</option>`).join("")}</select></label>
      <label>Horário do script (0 a 23h)<input type="number" min="0" max="23" id="au-hora" data-ak="hora" value="${esc(a.hora)}" ${dis}></label>
      <label>Link da planilha Google<input type="text" id="au-url" data-ak="planilhaUrl" value="${esc(a.planilhaUrl)}" ${dis} placeholder="https://docs.google.com/spreadsheets/d/..."></label>
      <label>Aba com o inventário<input type="text" id="au-aba" data-ak="aba" value="${esc(a.aba)}" ${dis}></label>
    </div>
    <details class="howto" ${temPlan?"":"open"}><summary>Como ativar (uma vez por publisher)</summary><ol class="passos">${passos}</ol></details>
    <div class="codehead"><b>Código.gs</b><button class="btn sm" data-copy="script">Copiar script</button></div>
    <pre class="code" id="code-script">${esc(gerarScript({...d, auto:a}))}</pre>
    <div class="codehead"><b>appsscript.json</b><button class="btn sm" data-copy="manifesto">Copiar manifesto</button></div>
    <pre class="code" id="code-manifesto" style="max-height:180px">${esc(api ? MANIFESTO : MANIFESTO_EMAIL)}</pre>
    <p class="note">Enquanto a leitura automática não estiver ligada, baixe a aba "${esc(a.aba)}" em CSV (Arquivo › Fazer download) e importe em Inventário: o formato já é o que o planejador espera.</p>
  </div>`;
}
function secaoMarca(d, ro){
  const m = {...MARCA_PADRAO(), ...(d.marca||{})}, pl = paleta(m.cor), url = logoUrl(m), dis = ro ? "disabled" : "";
  return `<div class="panel section"><h3>Identidade visual</h3>
    <p class="note">Logo e cor que os usuários de ${esc(d.nome)} veem no planejador e no Excel exportado. As demais cores são derivadas da cor principal, com contraste garantido para leitura.</p>
    <div class="brandprev">${url ? `<img src="${esc(url)}" alt="Logo de ${esc(d.nome)}">` : `<span class="logotext" style="font:700 18px var(--f-display)">${esc(d.nome)}</span>`}
      <div class="swatches" aria-label="Paleta derivada">${[pl.accent, pl.strong, pl.soft, pl.sunken, pl.ink].map(c=>`<span class="swatch" style="background:${c}" title="${c}"></span>`).join("")}</div>
      <span class="btn sm primary" style="pointer-events:none">Botão principal</span><span class="chip accent">Destaque</span></div>
    <div class="fields">
      <label>Cor principal<span class="colorrow"><input type="color" id="mk-cor" data-mk="cor" value="${esc(pl.accent.toLowerCase())}" ${dis} aria-label="Escolher cor"><input type="text" id="mk-hex" data-mk="cor" value="${esc(pl.accent)}" ${dis} maxlength="7" style="max-width:110px" aria-label="Cor em hexadecimal"></span><span class="hint">Use a cor da marca, como no logo.</span></label>
      <label>Logotipo${ro ? "" : S.assets ? `<span class="file"><label class="btn sm" for="mk-file" style="cursor:pointer">${url ? "Trocar logo" : "Enviar logo"}</label><input type="file" id="mk-file" accept="image/png,image/jpeg,image/webp,image/svg+xml" hidden>${url ? `<button class="btn sm" id="mk-rm">Remover logo</button>` : ""}</span><span class="hint">PNG ou SVG com fundo transparente, na horizontal. PNG ou JPG também entram no Excel.</span>` : `<span class="hint">O envio de arquivos não está disponível nesta visualização.</span>`}</label>
    </div>
    <div class="status" id="mk-msg"></div>
  </div>`;
}
function renderPublisher(){
  const box = $("#view-publisher");
  syncDraft();
  if(S.creating || (!curPub() && !S.loading)){
    box.innerHTML = S.canEdit && !S.dbOff ? `<div class="panel section" style="max-width:560px">
      <h3>Novo publisher</h3>
      <p class="note">O cadastro começa com os formatos padrão (display, formatos diferenciados e vídeo) e os CPMs de referência. Depois você ajusta e importa o relatório do GAM.</p>
      <div class="fields"><label>Nome<input type="text" id="np-nome" placeholder="Ex.: MarketUP"></label><label>Domínio<input type="text" id="np-dom" placeholder="ex.: marketup.com.br"></label></div>
      <div class="status bad" id="np-msg"></div>
      <div style="display:flex;gap:8px"><button class="btn primary" id="np-criar">Criar publisher</button>${curPub()?`<button class="btn" id="np-cancel">Cancelar</button>`:""}</div></div>`
      : `<div class="panel empty">Só administradores cadastram publishers.</div>`;
    const c = $("#np-criar"); if(c) c.onclick = criarPublisher;
    const x = $("#np-cancel"); if(x) x.onclick = () => { S.creating = false; renderPublisher(); };
    return;
  }
  const d = S.draft;
  if(!d){ box.innerHTML = `<div class="panel empty">Carregando…</div>`; return; }
  const ro = !(S.canEdit && S.db);
  const dis = ro ? "disabled" : "";
  const inv = d.inventario;
  const fmtRows = (d.formatos||[]).map((f,i) => `<tr>
    <td>${ro?"":`<input type="checkbox" id="fu-${i}" data-mrg="${i}" ${S.mergeSel.has(String(i))?"checked":""} aria-label="Selecionar ${esc(f.nome)} para unir">`}</td>
    <td><input type="checkbox" id="fa-${i}" data-fi="${i}" data-fk="ativo" ${f.ativo!==false?"checked":""} ${dis} aria-label="Ativo"></td>
    <td><input type="text" id="fn-${i}" data-fi="${i}" data-fk="nome" value="${esc(f.nome)}" ${dis} aria-label="Nome do formato"></td>
    <td><input class="n" type="number" step="0.5" min="0" id="fc-${i}" data-fi="${i}" data-fk="cpm" value="${esc(f.cpm)}" ${dis} aria-label="CPM tabela"></td>
    <td><input class="n" type="number" step="5" min="0" id="fm-${i}" data-fi="${i}" data-fk="mix" value="${esc(f.mix)}" ${dis} aria-label="Percentual padrão"></td>
    <td><input class="n" type="number" step="0.1" min="0" id="fp-${i}" data-fi="${i}" data-fk="impPorPv" value="${esc(f.impPorPv)}" ${dis} aria-label="Impressões por pageview"></td>
    <td><input type="text" id="fk-${i}" data-fi="${i}" data-fk="chaves" value="${esc((f.chaves||[]).join(", "))}" ${dis} aria-label="Palavras do nome da unidade de anúncio"></td>
    <td class="num">${inv && inv.fonte==="gam" && (f.id in (inv.porFormatoUF||{})) ? NUM.format(somaObj(inv.porFormatoUF[f.id])) : "—"}</td>
    <td><div class="rowact" style="justify-content:flex-start">${ro?"":`${(f.chaves||[]).length>1?`<button class="btn sm" data-splitfmt="${i}" title="Cria um formato para cada palavra da lista, cada um com o próprio CPM tabela">Separar por bloco</button>`:""}<button class="btn sm" data-rmfmt="${i}" aria-label="Remover ${esc(f.nome)}">Remover</button>`}</div></td></tr>`).join("");
  const vseg = volSegDia(d), vtot = vseg ? somaObj(vseg) : 0;
  const segRows = (d.segmentos||[]).map((s,i) => `<tr>
    <td><input class="n" type="text" id="si-${i}" data-si="${i}" data-sk="id" value="${esc(s.id)}" ${dis} aria-label="Valor no GAM" style="width:80px"></td>
    <td><input type="text" id="sn-${i}" data-si="${i}" data-sk="nome" value="${esc(s.nome)}" ${dis} aria-label="Nome do segmento"></td>
    <td><input type="text" id="sp-${i}" data-si="${i}" data-sk="palavras" value="${esc((s.palavras||[]).join(", "))}" ${dis} aria-label="Palavras-chave"></td>
    <td class="num">${vseg ? NUM.format(vseg[s.id]||0) : "—"}</td>
    <td class="num">${vseg && vtot ? pct((vseg[s.id]||0)/vtot) : "—"}</td>
    <td>${ro?"":`<button class="btn sm" data-rmseg="${i}" aria-label="Remover ${esc(s.nome)}">Remover</button>`}</td></tr>`).join("") +
    (vseg && vseg.ND ? `<tr><td></td><td class="note">Sem segmento</td><td></td><td class="num">${NUM.format(vseg.ND)}</td><td class="num">${pct(vseg.ND/vtot)}</td><td></td></tr>` : "");
  const ativosCfg = (d.formatos||[]).filter(f=>f.ativo!==false);
  const semInvFmts = inv && inv.fonte==="gam" ? ativosCfg.filter(f=>!(f.id in (inv.porFormatoUF||{}))).map(f=>f.nome) : [];
  const semCpm = ativosCfg.filter(f=>!(+f.cpm>0)).map(f=>f.nome);
  const invFmt = inv && inv.fonte==="gam" ? (d.formatos||[]).map(f => `<span>${esc(f.nome)}: <b>${NUM.format(somaObj(inv.porFormatoUF[f.id]))}</b>/dia</span>`).join("") : "";
  const topUF = inv && inv.fonte==="gam" ? (() => { const t = {}; for(const o of Object.values(inv.porFormatoUF)) for(const [uf,v] of Object.entries(o)) if(uf!=="ND") t[uf]=(t[uf]||0)+v; const tot = somaObj(t); return Object.entries(t).sort((a,b)=>b[1]-a[1]).slice(0,6).map(([uf,v])=>`<span>${uf}: <b>${pct(v/tot)}</b></span>`).join(""); })() : "";
  box.innerHTML = `<div class="pubview">
    <div class="result-head"><h2>${esc(d.nome)}</h2>${ro?`<span class="chip">Somente leitura</span>`:""}
      <div class="actions">${S.canEdit && !S.dbOff ? `<button class="btn" id="btn-novo-pub">Novo publisher</button>` : ""}</div></div>
    <div class="checklist">${passos(d).map((p,i)=>`<div class="check ${p.ok?"done":""}"><span class="dot">${p.ok?"✓":i+1}</span><div><b>${p.t}</b><span>${p.s}</span></div></div>`).join("")}</div>
    <div class="status ${S.pubMsg.startsWith("Não")?"bad":"ok"}" id="pub-msg">${esc(S.pubMsg)}</div>

    <div class="panel section"><h3>Dados do publisher</h3>
      <div class="fields">
        <label>Nome<input type="text" id="pd-nome" data-k="nome" value="${esc(d.nome)}" ${dis}></label>
        <label>Endereço do planejador<input type="text" value="${esc(enderecoDoPublisher(S.pubId).replace(/\/$/,""))}" readonly aria-label="Endereço do planejador deste publisher"><span class="hint">Subdomínio com o mesmo nome do código do publisher (${esc(S.pubId||"")}).</span></label>
        <label>Domínio<input type="text" id="pd-dom" data-k="dominio" value="${esc(d.dominio)}" ${dis} placeholder="ex.: marketup.com.br"></label>
        <label>Contato comercial<input type="text" id="pd-contato" data-k="contato" value="${esc(d.contato)}" ${dis}></label>
        <label>Código de rede do GAM<input type="text" id="pd-net" data-k="networkCode" value="${esc(d.networkCode)}" ${dis} inputmode="numeric"><span class="hint">Usado na integração automática pela API.</span></label>
      </div>
      <div class="fields">
        <label>Dedução sobre a verba (%)<input type="number" min="0" max="90" id="pd-ded" data-k="deducao" value="${esc(d.deducao)}" ${dis}><span class="hint">Agência e demais custos. Líquido = bruto − dedução.</span></label>
        <label>Teto de ocupação por campanha (%)<input type="number" min="0" max="100" id="pd-teto" data-k="tetoOcupacao" value="${esc(d.tetoOcupacao)}" ${dis}><span class="hint">Acima disso a linha é sinalizada.</span></label>
        <label>Acréscimo no CPM por estado (%)<input type="number" min="0" max="200" id="pd-geo" data-k="premioGeo" value="${esc(d.premioGeo)}" ${dis}><span class="hint">Aplicado quando a proposta é por estado.</span></label>
      </div>
    </div>

    ${secaoMarca(d, ro)}

    <div class="panel section"><h3>Segmentos do sistema</h3>
      <p class="note">Segmentos de público que o publisher vende. No GAM, cada um é um valor da chave-valor do segmento; o relatório traz o volume de impressões de cada um. As palavras-chave servem para sugerir segmentos a partir do contexto da campanha.</p>
      <div class="fields">
        <label>Chave do segmento no GAM<input type="text" id="pd-chave" data-k="chaveSegmento" value="${esc(d.chaveSegmento||"")}" ${dis} placeholder="ex.: segmento"><span class="hint">Nome da chave-valor. Com ela preenchida, outras chaves do relatório são descartadas.</span></label>
        <label>Acréscimo no CPM por segmento (%)<input type="number" min="0" max="200" id="pd-pseg" data-k="premioSeg" value="${esc(d.premioSeg||0)}" ${dis}><span class="hint">Aplicado quando a proposta escolhe segmentos.</span></label>
      </div>
      ${(d.segmentos||[]).length ? `<div class="tablebox"><table class="fmts"><thead><tr><th>Valor no GAM</th><th>Segmento</th><th>Palavras-chave no contexto</th><th class="num">Impressões/dia</th><th class="num">Participação</th><th></th></tr></thead><tbody>${segRows}</tbody></table></div>` : `<p class="note">Nenhum segmento cadastrado. Sem segmentos, a proposta usa só praça e formato.</p>`}
      ${ro?"":`<div><button class="btn sm" id="btn-addseg">Adicionar segmento</button></div>`}
    </div>

    <div class="panel section"><h3>Formatos e CPMs</h3>
      <p class="note">O CPM tabela é o preço de lista de cada formato; o desconto de cada campanha é dado na proposta. Cada formato reúne unidades de anúncio do GAM pelas palavras do nome delas. % padrão é a divisão inicial da verba na proposta; impressões por pageview só valem para a estimativa sem GAM.</p>
      <div class="tablebox"><table class="fmts"><thead><tr><th>Unir</th><th>Ativo</th><th>Formato</th><th class="num">CPM tabela R$</th><th class="num">% padrão</th><th class="num">Imp./PV</th><th>Palavras no nome da unidade de anúncio</th><th class="num">Impr./dia (GAM)</th><th></th></tr></thead><tbody>${fmtRows}</tbody></table></div>
      ${semInvFmts.length ? `<p class="note" style="color:var(--warn)">Sem inventário do GAM para: ${esc(semInvFmts.join(", "))}. Reimporte o relatório para medir esses formatos.</p>` : ""}
      ${semCpm.length ? `<p class="note" style="color:var(--bad)">Defina o CPM tabela de: ${esc(semCpm.join(", "))}.</p>` : ""}
      ${ro?"":`<div style="display:flex;gap:12px;align-items:center;flex-wrap:wrap"><button class="btn sm" id="btn-addfmt">Adicionar formato</button><button class="btn sm" id="btn-merge" title="Junta os formatos marcados na coluna Unir em um só, somando o inventário">Unir selecionados</button>${inv && inv.fonte==="gam" ? `<button class="btn sm" id="btn-mixinv" title="% padrão proporcional às impressões de cada formato">Distribuir % pelo inventário</button>` : ""}
        <label style="display:flex;gap:8px;align-items:center;font-size:13px;font-weight:600"><input type="checkbox" id="pd-cpmval" data-k="cpmValidado" ${d.cpmValidado?"checked":""}> CPMs validados com o comercial do publisher</label></div>`}
    </div>

    ${secaoAuto(d, ro)}

    <div class="panel section"><h3>Inventário (Google Ad Manager)</h3>
      ${inv && inv.fonte==="gam" ? `<div class="kv"><span class="chip ok">Dados reais do GAM</span><span>Arquivo: <b>${esc(inv.arquivo)}</b></span><span>Importado em <b>${esc(fmtData(inv.importadoEm))}</b></span><span>Média de <b>${esc(inv.dias)}</b> dias</span><span>Sem estado: <b>${pct(inv.semRegiaoPct||0)}</b></span></div>
        <div class="kv">${invFmt}</div><div class="kv"><span class="note">Principais estados:</span>${topUF}</div>` :
        `<p class="note">Sem relatório do GAM, a proposta calcula impressões pela verba e pelo CPM, mas não mostra a ocupação do inventário. ${+d.pageviewsMes>0?"Hoje a ocupação usa a estimativa por pageviews.":""}</p>`}
      ${ro?"":`<div class="file"><label class="btn primary" for="gam-file" style="cursor:pointer">Importar relatório do GAM (CSV ou Excel)</label><input type="file" id="gam-file" accept=".csv,.tsv,.txt,.xlsx,.xls" hidden></div>`}
      ${mostrarImportacao()}
      <details class="howto"><summary>Como gerar o relatório no GAM</summary><ol>
        <li>No Ad Manager do publisher, abra Relatórios e crie um novo relatório.</li>
        <li>Período: últimos 30 dias, ou o último mês fechado.</li>
        <li>Dimensões: Data, Unidade de anúncio, País e Região. Se o publisher tem segmentos, inclua também Chave-valor e filtre pela chave do segmento.</li>
        <li>Métrica: Total de impressões.</li>
        <li>Rode, exporte em CSV ou Excel e envie aqui. Salve o relatório no GAM para repetir todo mês.</li>
      </ol></details>
      <div class="fields"><label>Pageviews mensais (estimativa, sem GAM)<input type="number" min="0" step="10000" id="pd-pv" data-k="pageviewsMes" value="${esc(d.pageviewsMes||"")}" ${dis} placeholder="opcional"><span class="hint">Só é usado enquanto não houver relatório do GAM.</span></label></div>
    </div>
    ${S.dirty && !ro ? `<div class="savebar"><span class="note">Alterações não salvas</span><button class="btn" id="btn-descartar">Descartar</button><button class="btn primary" id="btn-salvar-pub">Salvar publisher</button></div>` : ""}
  </div>`;
  const on = (id, fn) => { const el = $(id); if(el) el.onclick = fn; };
  on("#btn-novo-pub", () => { S.creating = true; renderPublisher(); });
  on("#mk-rm", () => { d.marca = {...MARCA_PADRAO(), ...(d.marca||{}), logoUrl:"", logoAsset:"", logoArquivo:""}; marcarSujo(); renderPublisher(); render(); });
  const mf = $("#mk-file"); if(mf) mf.onchange = async e => {
    const file = e.target.files[0]; e.target.value = ""; if(!file) return;
    const msg = $("#mk-msg"); msg.className = "status"; msg.textContent = "Enviando logo…";
    const tipo = /\.svg$/i.test(file.name) ? "image/svg+xml" : /\.png$/i.test(file.name) ? "image/png" : /\.jpe?g$/i.test(file.name) ? "image/jpeg" : /\.webp$/i.test(file.name) ? "image/webp" : file.type;
    try{ const up = await S.assets.upload(file, {type:tipo}); d.marca = {...MARCA_PADRAO(), ...(d.marca||{}), logoUrl:up.url, logoAsset:"", logoArquivo:""}; marcarSujo(); renderPublisher(); render(); }
    catch(err){ msg.className = "status bad"; msg.textContent = err && err.code==="too_large" ? "O arquivo passou do limite de 2 MB." : err && err.code==="unsupported_type" ? "Formato não aceito. Envie PNG, JPG, WEBP ou SVG." : "Não foi possível enviar o logo agora. Tente de novo."; }
  };
  on("#btn-addseg", () => { (S.draft.segmentos ??= []).push({id:"", nome:"Novo segmento", palavras:[]}); marcarSujo(); renderPublisher(); });
  on("#btn-mixinv", () => { const tt = {}; for(const f of ativosCfg) tt[f.id] = somaObj((inv.porFormatoUF||{})[f.id]); const T = somaObj(tt); if(!T) return; for(const f of d.formatos) f.mix = f.ativo!==false ? Math.round((tt[f.id]||0)/T*1000)/10 : 0; marcarSujo(); renderPublisher(); });
  on("#btn-merge", () => {
    const idx = [...S.mergeSel].map(Number).filter(i=>d.formatos[i]).sort((a,b)=>a-b);
    if(idx.length < 2){ S.pubMsg = "Não foi possível unir: marque ao menos dois formatos na coluna Unir."; renderPublisher(); return; }
    const fs = idx.map(i=>d.formatos[i]);
    const palavras = fs.map(f=>String(f.nome).trim().split(/\s+/)); const comum = [];
    for(let k=0; palavras[0][k] && palavras.every(ws => ws[k] && norm(ws[k])===norm(palavras[0][k])); k++) comum.push(palavras[0][k]);
    const nome = comum.length ? comum.join(" ") : fs.map(f=>f.nome).sort((a,b)=>a.length-b.length)[0];
    let id = "blk-" + slug(nome); if(d.formatos.some((x,i)=>!idx.includes(i) && x.id===id)) id += "-" + Date.now().toString(36).slice(-4);
    const novo = {id, nome, cpm:+fs[0].cpm||0, mix:+fs.reduce((a,f)=>a+(+f.mix||0),0).toFixed(1), impPorPv:+fs.reduce((a,f)=>a+(+f.impPorPv||0),0).toFixed(2), ativo:true, chaves:[...new Set(fs.flatMap(f=>f.chaves||[]))]};
    const iv = d.inventario;
    if(iv) for(const key of ["porFormatoUF","porFormatoUFSeg"]){
      const o = iv[key]; if(!o) continue; const acc = {};
      for(const f of fs){ const src = o[f.id]; if(!src) continue;
        for(const uf in src){ if(key==="porFormatoUF") acc[uf] = (acc[uf]||0) + (+src[uf]||0); else { acc[uf] ??= {}; for(const sg in src[uf]) acc[uf][sg] = Math.round(((acc[uf][sg]||0) + (+src[uf][sg]||0))*10)/10; } }
        delete o[f.id]; }
      if(Object.keys(acc).length) o[id] = acc;
    }
    d.formatos.splice(idx[0], 1, novo); for(const i of idx.slice(1).reverse()) d.formatos.splice(i, 1);
    const cpmsDif = new Set(fs.map(f=>+f.cpm||0)).size > 1;
    S.mergeSel = new Set(); S.dirty = true;
    S.pubMsg = `${fs.length} formatos unidos em "${nome}".` + (cpmsDif ? ` Os CPMs eram diferentes; ficou o de ${fs[0].nome} (${BRL.format(+fs[0].cpm||0)}). Confira e salve.` : " Confira o nome e salve.");
    renderPublisher();
  });
  on("#btn-addfmt", () => { d.formatos.push({id:"fmt-"+Date.now().toString(36), nome:"Novo formato", cpm:15, mix:0, impPorPv:0.5, ativo:true, chaves:[]}); marcarSujo(); renderPublisher(); });
  on("#btn-descartar", () => { S.dirty = false; S.imp = null; syncDraft(); S.pubMsg=""; render(); });
  on("#btn-salvar-pub", () => salvarPublisher("Publisher salvo."));
  const fi = $("#gam-file"); if(fi) fi.onchange = e => { const file = e.target.files[0]; e.target.value = ""; if(file) iniciarImportacao(file); };
  const id = $("#imp-dias"); if(id) id.onchange = e => { S.imp.diasManual = Math.max(1, +e.target.value||30); agregar(); renderPublisher(); };
}
function marcarSujo(){
  if(S.dirty) return; S.dirty = true;
  if(!$(".savebar")) renderPublisherKeepFocus();
}
function renderPublisherKeepFocus(){
  const a = document.activeElement, id = a && a.id, s = a && a.selectionStart;
  renderPublisher();
  if(id){ const el = document.getElementById(id); if(el){ el.focus(); try{ if(s!=null) el.setSelectionRange(s,s); }catch(e){} } }
}
async function criarPublisher(){
  const nome = $("#np-nome").value.trim(), dom = $("#np-dom").value.trim(), msg = $("#np-msg");
  if(!nome){ msg.textContent = "Informe o nome do publisher."; return; }
  let id = slug(nome); if(S.publishers[id]) id += "-" + Date.now().toString(36).slice(-4);
  try{
    await S.db.collection("publishers").doc(id).set(novoPublisher(nome, dom));
    S.pubId = id; lsSet("pp:pub", id); S.creating = false; S.dirty = false; S.draftFor = null; S.pubMsg = "Publisher criado. Revise os formatos e importe o relatório do GAM.";
    render();
  } catch(e){ msg.textContent = "Não foi possível criar agora. Só administradores cadastram publishers."; }
}
async function sincronizarAcessos(pubId, dados){
  return {total:0, falhas:0};
  const alvo = S.acessos.filter(a=>a.publisherId===pubId); let falhas = 0;
  for(const a of alvo){ const {id, ...resto} = a; try{ await S.db.collection("acessos").doc(id).set({...resto, publisher:copiaPublica(dados), sincronizadoEm:new Date().toISOString()}); }catch(e){ falhas++; } }
  return {total:alvo.length, falhas};
}
async function salvarPublisher(okMsg){
  const d = S.draft; if(!d || !S.db) return;
  const limpo = clone(d); delete limpo.id;
  limpo.formatos = (limpo.formatos||[]).map(f => ({...f, cpm:+f.cpm||0, mix:+f.mix||0, impPorPv:+f.impPorPv||0}));
  limpo.segmentos = (limpo.segmentos||[]).filter(x=>String(x.id||"").trim()).map(x=>({id:String(x.id).trim(), nome:x.nome||"", palavras:x.palavras||[]}));
  for(const k of ["deducao","tetoOcupacao","premioGeo","premioSeg","pageviewsMes"]) limpo[k] = +limpo[k]||0;
  limpo.auto = {...AUTO_PADRAO(), ...(limpo.auto||{})};
  limpo.marca = {...MARCA_PADRAO(), ...(limpo.marca||{})}; limpo.marca.cor = hexOk(limpo.marca.cor) || "#2E7D4F";
  limpo.atualizadoEm = new Date().toISOString();
  try{
    await S.db.collection("publishers").doc(S.pubId).set(limpo); S.dirty = false; S.draftFor = null;
    const n = await sincronizarAcessos(S.pubId, limpo);
    S.pubMsg = okMsg + (n.total ? (n.falhas ? ` ${n.falhas} de ${n.total} usuários não foram atualizados; salve de novo.` : ` Acesso de ${n.total} usuário${n.total>1?"s":""} atualizado.`) : "");
  }
  catch(e){ S.pubMsg = "Não foi possível salvar o publisher. Só administradores podem alterar a configuração."; }
  renderPublisher(); renderMix();
}

/* ---------- usuários (admin) ---------- */
const enderecoSite = pid => enderecoDoPublisher(pid);
function textoConvite(email, pid, jaTemConta){
  const pub = (S.publishers[pid]||{}).nome || "seu publisher";
  return jaTemConta
    ? `Olá! Seu acesso ao Planejador de Propostas da ${pub} foi liberado.\n\nAcesse ${enderecoSite(pid)} e entre com o e-mail ${email} e a sua senha.`
    : `Olá! Você foi convidado para o Planejador de Propostas da ${pub}.\n\n1. Acesse ${enderecoSite(pid)}\n2. Clique em "Primeiro acesso? Criar conta"\n3. Use o e-mail ${email} e crie a sua senha\n\nVocê entra direto no planejador da ${pub}.`;
}
const optsPub = sel => Object.keys(S.publishers).sort((a,b)=>String(S.publishers[a].nome).localeCompare(String(S.publishers[b].nome))).map(id=>`<option value="${esc(id)}"${id===sel?" selected":""}>${esc(S.publishers[id].nome)}</option>`).join("");
async function perfis(ids){ if(!ids.length) return {}; return DATA.perfis(ids); }
const pessoa = (p, extra) => `<div class="person"><img alt="" src="${esc(p && p.avatarUrl || "")}"><div><b>${esc((p && p.name) || "Usuário")}</b><span>${esc(extra || "")}</span></div></div>`;
async function concederAcesso(uid, pid, empresa){
  const pub = S.publishers[pid]; if(!uid || !pub) return false;
  try{ await S.db.collection("acessos").doc(uid).set({publisherId:pid}); return true; }
  catch(e){ return false; }
}
function pendentes(){
  const com = new Set(S.acessos.map(a=>a.id));
  return S.perfis.filter(p => !p.admin && !com.has(p.id)).map(p => ({uid:p.id, empresa:p.empresa||"", email:p.email||""}));
}
async function renderPendentes(){
  const box = $("#u-pend"); if(!box) return;
  const ps = pendentes();
  if(!ps.length){ box.innerHTML = `<p class="note">Ninguém aguardando. Quando a pessoa entrar com o e-mail e pedir acesso, ela aparece aqui.</p>`; return; }
  const pf = await perfis(ps.map(x=>x.uid));
  box.innerHTML = `<div class="ures">${ps.map(x => `<div class="item pend">${pessoa(pf[x.uid], `${x.email}${x.empresa ? " · Empresa informada: "+x.empresa : " · empresa não informada"}`)}
    <span style="display:flex;gap:6px;flex-wrap:wrap"><select class="in" id="pd-${esc(x.uid)}" style="width:auto">${optsPub(S.pubId)}</select><button class="btn sm primary" data-conceder="${esc(x.uid)}" data-emp="${esc(x.empresa)}">Liberar acesso</button></span></div>`).join("")}</div>`;
}
async function renderUsuarios(){
  const box = $("#view-usuarios");
  if(!ehAdmin()){ box.innerHTML = ""; return; }
  const nomePub = a => (S.publishers[a.publisherId]||{}).nome || "";
  const lista = S.acessos.slice().sort((a,b)=>nomePub(a).localeCompare(nomePub(b)));
  const pf = await perfis(lista.map(a=>a.id));
  const ativo = document.activeElement && document.activeElement.id;
  box.innerHTML = `<div class="pubview">
    <div class="result-head"><h2>Usuários e acessos</h2></div>
    <div class="status ${S.uMsg.startsWith("Não")?"bad":"ok"}">${esc(S.uMsg)}</div>
    <div class="panel section"><h3>Convidar colaborador</h3>
      <p class="note">Registre o e-mail e o publisher e envie a mensagem gerada (WhatsApp ou e-mail). A pessoa cria a senha com esse e-mail e entra direto no publisher, sem precisar pedir liberação.</p>
      <form class="fields" id="conv-form" novalidate>
        <label>E-mail do colaborador<input type="email" id="conv-email" autocomplete="off" placeholder="nome@empresa.com.br" value="${esc(S.convEmail||"")}"></label>
        <label>Publisher<select class="in" id="conv-pub">${optsPub(S.convPub || S.pubId)}</select></label>
        <label style="justify-content:flex-end"><button class="btn primary" type="submit" id="btn-convidar">Registrar convite</button></label>
      </form>
      ${S.convMsg ? `<div><div class="codehead"><b>Mensagem para enviar</b><button class="btn sm" data-copiarconv="1">Copiar mensagem</button></div><pre class="convmsg" id="conv-texto">${esc(S.convMsg)}</pre></div>` : ""}
      ${S.convites.length ? `<div><b>Convites aguardando o primeiro acesso</b><div class="tablebox" style="margin-top:8px"><table class="users"><thead><tr><th>E-mail</th><th>Publisher</th><th>Enviado em</th><th class="num">Ações</th></tr></thead><tbody>${S.convites.map(c => `<tr>
        <td>${esc(c.email)}</td><td>${esc((S.publishers[c.publisher_id]||{}).nome || c.publisher_id)}</td><td>${esc(fmtData(c.criado_em))}</td>
        <td class="num"><div class="rowact"><button class="btn sm" data-reenviar="${esc(c.email)}" data-pub="${esc(c.publisher_id)}">Copiar mensagem</button>${S.confirmConv===c.email ? `<button class="btn sm danger" data-cancconv="${esc(c.email)}" data-confirm="1">Confirmar cancelamento</button>` : `<button class="btn sm" data-cancconv="${esc(c.email)}">Cancelar convite</button>`}</div></td></tr>`).join("")}</tbody></table></div></div>` : ""}
    </div>
    <div class="panel section"><h3>Pedidos de acesso</h3>
      <p class="note">Quem entra no planejador sem convite pode pedir acesso informando a empresa. Os pedidos aparecem em <b>Aguardando acesso</b>: escolha o publisher e clique em <b>Liberar acesso</b>. Cada usuário lê só o publisher liberado para ele; o banco recusa qualquer outra leitura.</p>
    </div>
    <div class="panel section"><h3>Aguardando acesso</h3><div id="u-pend"><p class="note">Carregando…</p></div></div>
    <div class="panel section"><h3>Usuários com acesso <span class="chip">${lista.length}</span></h3>
      ${lista.length ? `<div class="tablebox"><table class="users"><thead><tr><th>Pessoa</th><th>Publisher</th><th>Liberado em</th><th class="num">Ações</th></tr></thead><tbody>${lista.map(a => `<tr>
        <td>${pessoa(pf[a.id], [pf[a.id] && pf[a.id].email, pf[a.id] && pf[a.id].empresa ? "Empresa informada: "+pf[a.id].empresa : ""].filter(Boolean).join(" · "))}</td>
        <td><select class="in" data-trocapub="${esc(a.id)}" style="width:auto" aria-label="Publisher do usuário">${optsPub(a.publisherId)}</select></td>
        <td>${esc(fmtData(a.concedidoEm))}</td>
        <td class="num"><div class="rowact">${S.confirmRm===a.id ? `<button class="btn sm danger" data-rmacesso="${esc(a.id)}" data-confirm="1">Confirmar remoção</button>` : `<button class="btn sm" data-rmacesso="${esc(a.id)}">Remover acesso</button>`}</div></td></tr>`).join("")}</tbody></table></div>`
        : `<p class="note">Nenhum usuário de publisher ainda.</p>`}
      <p class="note">Remover o acesso tira a leitura dos dados na hora; a pessoa volta para a tela de pedido de acesso.</p>
    </div>
  </div>`;
  if(ativo){ const el = document.getElementById(ativo); if(el) el.focus(); }
  renderPendentes();
}
async function buscarPessoas(q){
  const box = $("#u-res"); if(!box || !S.user) return;
  const hits = await S.user.search(q);
  const com = new Set(S.acessos.map(a=>a.id));
  const lista = hits.filter(h=>!h.isMe);
  box.innerHTML = !q ? "" : !lista.length ? `<p class="note">Nenhuma pessoa encontrada.</p>` : lista.map(h => `<div class="item">${pessoa(h, h.email || "")}
    ${com.has(h.id) ? `<span class="chip ok">Já tem acesso</span>` : `<span style="display:flex;gap:6px;flex-wrap:wrap"><select class="in" id="pd-${esc(h.id)}" style="width:auto">${optsPub(S.pubId)}</select><button class="btn sm primary" data-conceder="${esc(h.id)}" data-emp="">Liberar acesso</button></span>`}</div>`).join("");
}

/* ---------- eventos ---------- */
document.addEventListener("click", e => {
  const t = e.target.closest("button, [data-go]"); if(!t) return;
  if(t.dataset.view){ S.view = t.dataset.view; S.confirmDel = null; S.confirmRm = null; render(); return; }
  if(t.id==="btn-preview"){ S.preview = true; S.view = "planejar"; render(); return; }
  if(t.id==="btn-preview-sair"){ S.preview = false; render(); return; }
  if(t.id==="btn-sair"){ DATA.sair(); return; }
  if(t.id==="btn-senha"){ S.trocaSenha = "trocar"; render(); return; }
  if(t.dataset.conceder){
    const uid = t.dataset.conceder, sel = document.getElementById("pd-"+uid), pid = sel && sel.value;
    t.disabled = true;
    concederAcesso(uid, pid, t.dataset.emp).then(ok => { S.uMsg = ok ? `Acesso a ${S.publishers[pid].nome} liberado.` : "Não foi possível liberar o acesso agora."; renderUsuarios(); });
    return;
  }
  if(t.dataset.reenviar){ S.convMsg = textoConvite(t.dataset.reenviar, t.dataset.pub); S.uMsg = ""; renderUsuarios(); return; }
  if(t.dataset.copiarconv){
    const pre = $("#conv-texto"), txt = pre.textContent;
    const ok = () => { t.textContent = "Copiado"; setTimeout(()=>{ t.textContent = "Copiar mensagem"; }, 1500); };
    const fb = () => { const sel = window.getSelection(), rg = document.createRange(); rg.selectNodeContents(pre); sel.removeAllRanges(); sel.addRange(rg); t.textContent = "Selecionado: use Ctrl+C"; };
    try{ navigator.clipboard.writeText(txt).then(ok, fb); }catch(err){ fb(); }
    return;
  }
  if(t.dataset.cancconv){
    if(!t.dataset.confirm){ S.confirmConv = t.dataset.cancconv; renderUsuarios(); return; }
    DATA.cancelarConvite(t.dataset.cancconv).then(ok => { S.uMsg = ok ? "Convite cancelado." : "Não foi possível cancelar agora."; S.confirmConv = null; renderUsuarios(); });
    return;
  }
  if(t.dataset.rmacesso){
    if(!t.dataset.confirm){ S.confirmRm = t.dataset.rmacesso; renderUsuarios(); return; }
    S.db.collection("acessos").doc(t.dataset.rmacesso).delete().then(() => { S.uMsg = "Acesso removido."; }, () => { S.uMsg = "Não foi possível remover o acesso agora."; }).finally(() => { S.confirmRm = null; renderUsuarios(); });
    return;
  }
  if(t.dataset.go){ S.view = t.dataset.go; render(); return; }
  if(t.dataset.uf){ S.ufs.has(t.dataset.uf) ? S.ufs.delete(t.dataset.uf) : S.ufs.add(t.dataset.uf); setUFs([...S.ufs]); atualizarHints(); render(); return; }
  if(t.dataset.copy){
    const pre = $(t.dataset.copy==="script" ? "#code-script" : "#code-manifesto"), txt = pre.textContent;
    const ok = () => { const o = t.textContent; t.textContent = "Copiado"; setTimeout(()=>{ t.textContent = o; }, 1500); };
    const fallback = () => { const sel = window.getSelection(), rg = document.createRange(); rg.selectNodeContents(pre); sel.removeAllRanges(); sel.addRange(rg); t.textContent = "Selecionado: use Ctrl+C"; };
    try{ navigator.clipboard.writeText(txt).then(ok, fallback); }catch(err){ fallback(); }
    return;
  }
  if(t.dataset.fall){ S.fmtsSel = t.dataset.fall==="todos" ? null : new Set(); renderFmtSel(); renderResult(); return; }
  if(t.dataset.reg){ if(t.dataset.reg==="limpar") setUFs([]); else { const r = UFS.filter(u=>u[3]===t.dataset.reg).map(u=>u[0]); const todos = r.every(u=>S.ufs.has(u)); setUFs(todos ? [...S.ufs].filter(u=>!r.includes(u)) : [...S.ufs, ...r]); } atualizarHints(); render(); return; }
  if(t.dataset.filter){ S.filtro = t.dataset.filter; renderSalvos(); return; }
  if(t.dataset.abrir){ const p = S.planos.find(x=>x.id===t.dataset.abrir); if(p) carregarPlano(p, true); return; }
  if(t.dataset.ppt){ const p = S.planos.find(x=>x.id===t.dataset.ppt); if(p) exportarPlano(p, "ppt"); return; }
  if(t.dataset.exp){ const p = S.planos.find(x=>x.id===t.dataset.exp); if(p) exportarPlano(p); return; }
  if(t.dataset.st){ setStatus(t.dataset.st, t.dataset.to); return; }
  if(t.dataset.del){ if(t.dataset.confirm) excluir(t.dataset.del); else { S.confirmDel = t.dataset.del; renderSalvos(); } return; }
  if(t.dataset.rmseg!=null){ S.draft.segmentos.splice(+t.dataset.rmseg,1); S.dirty = true; renderPublisher(); return; }
  if(t.dataset.splitfmt!=null || t.dataset.rmfmt!=null || t.id==="btn-addfmt" || t.id==="btn-descartar") S.mergeSel = new Set();
  if(t.dataset.splitfmt!=null){
    const i = +t.dataset.splitfmt, f = S.draft.formatos[i], ks = f.chaves||[];
    const novos = ks.map(k => ({id:"blk-"+slug(k), nome:nomeBloco(k), cpm:+f.cpm||0, mix:+((+f.mix||0)/ks.length).toFixed(1), impPorPv:+((+f.impPorPv||0)/ks.length).toFixed(2), ativo:f.ativo!==false, chaves:[k]}))
      .filter((x,j,a)=>a.findIndex(y=>y.id===x.id)===j && !S.draft.formatos.some((y,yi)=>yi!==i && y.id===x.id));
    S.draft.formatos.splice(i, 1, ...novos); S.draft.cpmValidado = false; S.dirty = true;
    S.pubMsg = `${f.nome} separado em ${novos.length} formatos. Ajuste o CPM tabela de cada um, salve e reimporte o relatório do GAM.`;
    renderPublisher(); return;
  }
  if(t.dataset.rmfmt!=null){ S.draft.formatos.splice(+t.dataset.rmfmt,1); S.dirty = true; renderPublisher(); return; }
  if(t.dataset.imp==="cancelar"){ S.imp = null; renderPublisher(); return; }
  if(t.dataset.imp==="confirmar"){ confirmarImportacao(); return; }
});
document.addEventListener("input", e => {
  const t = e.target;
  if(t.closest("#form")){
    if(t.name==="segmode" || t.dataset.seg!=null || t.dataset.fsel!=null) return;
    if(t.dataset.mix){ S.mix[t.dataset.mix] = t.value===""?0:+t.value; atualizarHintMix(); }
    if(t.dataset.cpm){ S.cpm[t.dataset.cpm] = t.value===""?0:+t.value; }
    S.saveMsg = ""; atualizarHints(); renderResult(); return;
  }
  if(t.id==="f-busca"){ S.busca = t.value; renderSalvos(); return; }
  if(t.id==="conv-email"){ S.convEmail = t.value; return; }
  if(t.id==="u-busca"){ S.uBusca = t.value.trim(); buscarPessoas(S.uBusca); return; }
  if(S.draft && t.closest("#view-publisher")){
    if(t.name==="automodo") return;
    if(t.dataset.mk==="cor"){ const v = t.type==="color" ? t.value : t.value.trim(); const h = hexOk(v.startsWith("#") ? v : "#"+v); if(!h) return;
      S.draft.marca = {...MARCA_PADRAO(), ...(S.draft.marca||{}), cor:h}; S.dirty = true;
      const other = t.type==="color" ? $("#mk-hex") : $("#mk-cor"); if(other) other.value = t.type==="color" ? h : h.toLowerCase();
      aplicarMarca(S.draft); return; }
    if(t.dataset.ak){ S.draft.auto ??= AUTO_PADRAO(); S.draft.auto[t.dataset.ak] = t.dataset.ak==="hora" ? Math.min(23, Math.max(0, parseInt(t.value,10)||0)) : t.value.trim(); marcarSujo(); atualizaCodigo(); return; }
    if(t.dataset.k){ const k = t.dataset.k; S.draft[k] = t.type==="checkbox" ? t.checked : (t.type==="number" ? (t.value===""?0:+t.value) : t.value); marcarSujo(); return; }
    if(t.dataset.si!=null){ const sg = S.draft.segmentos[+t.dataset.si], k = t.dataset.sk;
      sg[k] = k==="palavras" ? t.value.split(",").map(s=>s.trim()).filter(Boolean) : t.value.trim();
      marcarSujo(); return; }
    if(t.dataset.fi!=null){ const f = S.draft.formatos[+t.dataset.fi], k = t.dataset.fk;
      f[k] = k==="ativo" ? t.checked : k==="chaves" ? t.value.split(",").map(s=>s.trim()).filter(Boolean) : (t.type==="number" ? (t.value===""?0:+t.value) : t.value);
      marcarSujo(); return; }
  }
});
function atualizaCodigo(){ const el = $("#code-script"); if(el && S.draft) el.textContent = gerarScript({...S.draft, auto:{...AUTO_PADRAO(), ...(S.draft.auto||{})}}); }
document.addEventListener("change", e => {
  const t = e.target;
  if(t.dataset.mk==="cor"){ renderPublisherKeepFocus(); return; }
  if(t.dataset.ak){ S.draft.auto ??= AUTO_PADRAO(); S.draft.auto[t.dataset.ak] = t.value; marcarSujo(); renderPublisher(); return; }
  if(t.name==="automodo" && S.draft){ S.draft.auto = {...AUTO_PADRAO(), ...(S.draft.auto||{}), modo:t.value}; S.dirty = true; renderPublisher(); return; }
  if(t.name==="praca"){ atualizarHints(); renderResult(); return; }
  if(t.name==="segmode"){ S.modoSeg = t.value; renderSegs(); atualizarHints(); renderResult(); return; }
  if(t.dataset.fsel!=null){ const pub = curPub(); if(!S.fmtsSel) S.fmtsSel = new Set(fmtsSelecionados(pub)); t.checked ? S.fmtsSel.add(t.dataset.fsel) : S.fmtsSel.delete(t.dataset.fsel); atualizarHintFsel(); renderResult(); return; }
  if(t.dataset.seg!=null && t.closest("#form")){ const id = String(t.dataset.seg); t.checked ? S.segs.add(id) : S.segs.delete(id); atualizarHintSegs(); sugerirSegs(); renderResult(); return; }
  if(t.id==="f-verba"){ const v = parseMoney(t.value); t.value = v ? BRL.format(v) : ""; return; }
  if(t.id==="pub-select"){
    if(t.value==="__novo"){ S.creating = true; S.view = "publisher"; render(); return; }
    S.pubId = t.value; lsSet("pp:pub", S.pubId); S.mix = {}; S.cpm = {}; S.segs = new Set(); S.modoSeg = "todos"; S.fmtsSel = null; S.imp = null; S.dirty = false; S.creating = false; S.pubMsg = ""; S.saveMsg = "";
    renderMix(); atualizarHints(); render(); return;
  }
  if(t.dataset.trocapub){
    const a = S.acessos.find(x=>x.id===t.dataset.trocapub); if(!a) return;
    concederAcesso(a.id, t.value, a.empresaDeclarada).then(ok => { S.uMsg = ok ? `Acesso alterado para ${S.publishers[t.value].nome}.` : "Não foi possível alterar o acesso agora."; renderUsuarios(); });
    return;
  }
  if(t.dataset.mrg!=null){ t.checked ? S.mergeSel.add(t.dataset.mrg) : S.mergeSel.delete(t.dataset.mrg); return; }
  if(t.dataset.col && S.imp){ S.imp.cols[t.dataset.col] = +t.value; agregar(); renderPublisher(); return; }
  if(t.dataset.atrib!=null && S.imp){ S.imp.atrib[t.dataset.atrib] = t.value || undefined; if(!t.value) delete S.imp.atrib[t.dataset.atrib]; agregar(); renderPublisher(); return; }
});

document.addEventListener("submit", e => {
  e.preventDefault();
  if(e.target.id==="conv-form"){
    const btn = $("#btn-convidar"), email = $("#conv-email").value, pid = $("#conv-pub").value;
    S.convEmail = email; S.convPub = pid; btn.disabled = true; btn.textContent = "Enviando…";
    DATA.convidar(email, pid).then(r => { S.uMsg = r.ok ? r.msg : "Não foi possível convidar: " + r.msg; if(r.ok){ S.convMsg = textoConvite(r.email, pid, r.jaTemConta); S.convEmail = ""; } renderUsuarios(); });
  }
});

/* ---------- camada de dados: Supabase (login, banco, logos e downloads) ---------- */
const CFG = window.SMARTREE_CONFIG || {};
/* subdomínio do publisher: marketup.smartree.com.br → "marketup" (ou ?p=marketup para testes) */
const DOMINIO_BASE = String(CFG.dominioBase || "").trim().toLowerCase().replace(/^\.+|\/+$/g, "");
const SUBDOMINIO = (() => {
  const reservados = ["app","www","admin","api"];
  const h = location.hostname.toLowerCase();
  if(DOMINIO_BASE && h.endsWith("." + DOMINIO_BASE)){ const sub = h.slice(0, -DOMINIO_BASE.length - 1); if(/^[a-z0-9-]{1,64}$/.test(sub) && !reservados.includes(sub)) return sub; }
  const q = new URLSearchParams(location.search).get("p");
  return q && /^[a-z0-9-]{1,64}$/.test(q) ? q : null;
})();
const enderecoDoPublisher = pid => DOMINIO_BASE && /^[a-z0-9-]{1,64}$/.test(String(pid||"")) ? `https://${pid}.${DOMINIO_BASE}/` : location.origin + location.pathname;
/* aceita o endereço colado com caminho (ex.: .../rest/v1/) e usa só a origem */
const SB_URL = (String(CFG.supabaseUrl||"").trim().match(/^https:\/\/[a-z0-9-]+\.supabase\.co/i) || [""])[0];
const SB_KEY = String(CFG.supabaseAnonKey||"").trim();
const configOk = !!(window.supabase && SB_URL && SB_KEY.length > 20);
const sb = configOk ? window.supabase.createClient(SB_URL, SB_KEY, {auth:{persistSession:true, autoRefreshToken:true, detectSessionInUrl:true, flowType:"implicit"}}) : null;
const CORES_AVATAR = ["#2E7D4F","#5A2ED8","#9A6414","#B4313B","#1F6FB2","#7A4E2D"];
function avatarDe(texto){
  const t = String(texto||"?").trim(), ini = (t.split(/[\s@._-]+/).filter(Boolean).slice(0,2).map(p=>p[0]).join("") || "?").toUpperCase();
  let h = 0; for(const c of t) h = (h*31 + c.charCodeAt(0)) >>> 0;
  return "data:image/svg+xml," + encodeURIComponent(`<svg xmlns="http://www.w3.org/2000/svg" width="40" height="40"><rect width="40" height="40" rx="20" fill="${CORES_AVATAR[h % CORES_AVATAR.length]}"/><text x="20" y="25" font-family="Arial" font-size="15" fill="#fff" text-anchor="middle">${ini.replace(/[<>&"]/g,"")}</text></svg>`);
}
const perfilVisivel = p => ({id:p.id, name:p.nome || String(p.email||"").split("@")[0] || "Usuário", email:p.email||"", empresa:p.empresa||"", admin:!!p.admin, avatarUrl:avatarDe(p.nome || p.email)});
const erroDe = r => { if(r && r.error) throw r.error; return r; };

async function recarregarPublishers(){
  const r = await sb.from("publishers").select("id,dados"); if(r.error) return;
  const prev = S.pubId, salvo = lsGet("pp:pub");
  S.publishers = {}; for(const row of r.data) S.publishers[row.id] = {...row.dados, id:row.id};
  if(!S.publishers[S.pubId]) S.pubId = (salvo && S.publishers[salvo]) ? salvo : (Object.keys(S.publishers).sort()[0] || null);
  if(prev !== S.pubId){ S.mix = {}; S.cpm = {}; S.segs = new Set(); S.modoSeg = "todos"; S.fmtsSel = null; }
  if(S.view!=="publisher" || !S.dirty) renderMix();
  atualizarHints(); render();
}
async function recarregarPropostas(){
  const r = await sb.from("propostas").select("id,dados,criado_em").order("criado_em", {ascending:false}); if(r.error) return;
  S.planos = r.data.map(row => ({...row.dados, id:row.id}));
  bannerAprovado(); renderHeader(); if(S.view==="salvos") renderSalvos();
}
async function recarregarAcessos(){
  if(S.role!=="admin") return;
  const [a, p] = await Promise.all([sb.from("acessos").select("user_id,publisher_id,concedido_em"), sb.from("perfis").select("id,email,nome,empresa,admin,criado_em").order("criado_em", {ascending:false})]);
  if(!a.error) S.acessos = a.data.map(row => ({id:row.user_id, publisherId:row.publisher_id, concedidoEm:row.concedido_em}));
  if(!p.error) S.perfis = p.data;
  const c = await sb.from("convites").select("email,publisher_id,criado_em").order("criado_em", {ascending:false});
  S.convites = c.error ? [] : c.data;
  if(S.view==="usuarios") renderUsuarios();
}
const TABELA = {planos:"propostas", publishers:"publishers", acessos:"acessos"};
S.db = sb && { collection(nome){ return {
  async add(dados){
    if(nome!=="planos") throw new Error("operação não suportada");
    erroDe(await sb.from("propostas").insert({publisher_id:dados.publisherId, dados}));
    await recarregarPropostas();
  },
  doc(id){ return {
    async set(v){
      if(nome==="publishers"){ const {id:_, ...dados} = v; erroDe(await sb.from("publishers").upsert({id, dados, atualizado_em:new Date().toISOString()})); await recarregarPublishers(); return; }
      if(nome==="acessos"){ erroDe(await sb.from("acessos").upsert({user_id:id, publisher_id:v.publisherId, concedido_em:new Date().toISOString()})); await recarregarAcessos(); return; }
      throw new Error("operação não suportada");
    },
    async update(patch){
      if(nome!=="planos") throw new Error("operação não suportada");
      const atual = S.planos.find(p=>p.id===id); if(!atual) throw new Error("proposta não encontrada");
      const {id:_, ...dados} = {...atual, ...patch};
      erroDe(await sb.from("propostas").update({dados}).eq("id", id));
      await recarregarPropostas();
    },
    async delete(){
      const tab = TABELA[nome]; if(!tab) throw new Error("operação não suportada");
      erroDe(await sb.from(tab).delete().eq(nome==="acessos" ? "user_id" : "id", id));
      if(nome==="planos") await recarregarPropostas(); else if(nome==="acessos") await recarregarAcessos(); else await recarregarPublishers();
    }
  }; }
}; } };
S.dl = { async save({filename, data}){
  const tipo = /\.pptx$/i.test(filename) ? "application/vnd.openxmlformats-officedocument.presentationml.presentation" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  const nome = String(filename).normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[–—]/g, "-").replace(/[^\x20-\x7E]/g, "").replace(/\s+/g, " ").trim() || "proposta";
  const a = document.createElement("a"); a.href = URL.createObjectURL(new Blob([data], {type:tipo})); a.download = nome;
  document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1500);
  return {status:"saved"};
} };
const assetsSupabase = { async upload(file, {type} = {}){
  if(file.size > 2*1024*1024) throw {code:"too_large"};
  const ext = {"image/png":"png","image/jpeg":"jpg","image/webp":"webp","image/svg+xml":"svg"}[type || file.type]; if(!ext) throw {code:"unsupported_type"};
  const caminho = `${S.pubId || "geral"}/${Date.now()}.${ext}`;
  erroDe(await sb.storage.from("logos").upload(caminho, file, {contentType:type || file.type, upsert:false}));
  return {id:"", url:sb.storage.from("logos").getPublicUrl(caminho).data.publicUrl};
} };

let carregando = null;
async function carregarSessao(){
  if(carregando) return carregando;
  carregando = (async () => {
    const {data:{session}} = await sb.auth.getSession();
    if(!session){ Object.assign(S, {me:null, perfil:null, role:"login", loading:false, publishers:{}, pubId:null, planos:[], acessos:[], perfis:[], canEdit:false, assets:null}); render(); return; }
    const u = session.user;
    const pr = await sb.from("perfis").select("id,email,nome,empresa,admin").eq("id", u.id).maybeSingle();
    const perfil = pr.data || {id:u.id, email:u.email};
    S.perfil = perfil; S.me = {id:u.id, email:u.email, name:perfil.nome || String(u.email||"").split("@")[0], avatarUrl:avatarDe(perfil.nome || u.email)};
    if(perfil.admin){
      Object.assign(S, {role:"admin", canEdit:true, canWrite:true, assets:assetsSupabase});
      if(SUBDOMINIO) S.pubId = SUBDOMINIO;
      await recarregarPublishers(); await Promise.all([recarregarPropostas(), recarregarAcessos()]);
    } else {
      const ac = await sb.from("acessos").select("publisher_id").eq("user_id", u.id).maybeSingle();
      if(!ac.data){ Object.assign(S, {role:"sem-acesso", publishers:{}, pubId:null, planos:[], canEdit:false, assets:null}); }
      else {
        if(SUBDOMINIO && DOMINIO_BASE && location.hostname.toLowerCase().endsWith("." + DOMINIO_BASE) && SUBDOMINIO !== ac.data.publisher_id){
          Object.assign(S, {role:"outro-publisher", publishers:{}, pubId:null, planos:[], canEdit:false, assets:null, pubCerto:ac.data.publisher_id});
        } else {
          Object.assign(S, {role:"publisher", canEdit:false, canWrite:true, assets:null, preview:false, pubId:ac.data.publisher_id, view:"planejar"});
          await recarregarPublishers(); S.pubId = ac.data.publisher_id; await recarregarPropostas();
        }
      }
    }
    S.loading = false; renderMix(); atualizarHints(); render();
  })();
  try{ await carregando; } finally { carregando = null; }
}
const DATA = {
  async iniciar(){
    if(!sb){ S.role = "config"; S.loading = false; render(); return; }
    if(SUBDOMINIO){ try{ const r = await sb.rpc("marca_publica", {p_id:SUBDOMINIO}); if(!r.error && r.data && r.data.nome) S.marcaPublica = r.data; }catch(e){} }
    S.role = "carregando"; render();
    sb.auth.onAuthStateChange(ev => {
      if(ev==="PASSWORD_RECOVERY"){ S.trocaSenha = "recuperar"; setTimeout(() => carregarSessao().then(render), 0); return; }
      if(ev==="SIGNED_IN" || ev==="SIGNED_OUT") setTimeout(carregarSessao, 0);
    });
    await carregarSessao();
  },
  carregarSessao,
  async convidar(email, pid){
    email = String(email||"").trim().toLowerCase();
    if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return {ok:false, msg:"Informe um e-mail válido."};
    if(!S.publishers[pid]) return {ok:false, msg:"Escolha o publisher."};
    const existente = S.perfis.find(p => String(p.email||"").toLowerCase()===email);
    if(existente && existente.admin) return {ok:false, msg:"Esse e-mail já é de um administrador."};
    if(existente){
      const r = await sb.from("acessos").upsert({user_id:existente.id, publisher_id:pid, concedido_em:new Date().toISOString()});
      if(r.error) return {ok:false, msg:"Não foi possível liberar o acesso agora."};
      await recarregarAcessos();
      return {ok:true, email, jaTemConta:true, msg:`${email} já tinha conta: acesso a ${S.publishers[pid].nome} liberado. Envie a mensagem abaixo para avisar.`};
    }
    const r = await sb.from("convites").upsert({email, publisher_id:pid, criado_em:new Date().toISOString()});
    if(r.error) return {ok:false, msg:"Não foi possível registrar o convite. Confira se o trecho de SQL dos convites foi rodado no Supabase."};
    await recarregarAcessos();
    return {ok:true, email, jaTemConta:false, msg:`Convite registrado para ${email}. Envie a mensagem abaixo: ao criar a conta com esse e-mail, a pessoa entra direto em ${S.publishers[pid].nome}.`};
  },
  async cancelarConvite(email){
    const r = await sb.from("convites").delete().eq("email", email); await recarregarAcessos(); return !r.error;
  },
  async entrar(email, senha){
    const {error} = await sb.auth.signInWithPassword({email, password:senha});
    return error ? (error.message || "erro") : null;
  },
  async criarConta(email, senha){
    const {data, error} = await sb.auth.signUp({email, password:senha, options:{emailRedirectTo: location.origin + location.pathname}});
    if(error) return {erro: error.message || "erro"};
    if(data && data.user && Array.isArray(data.user.identities) && data.user.identities.length===0) return {erro:"already registered"};
    return {confirmar: !data.session};
  },
  async esqueciSenha(email){
    const {error} = await sb.auth.resetPasswordForEmail(email, {redirectTo: location.origin + location.pathname});
    return error ? (error.message || "erro") : null;
  },
  async definirSenha(senha){
    const {error} = await sb.auth.updateUser({password:senha});
    return error ? (error.message || "erro") : null;
  },
  async sair(){ await sb.auth.signOut(); S.loginOk = ""; S.loginModo = "entrar"; S.trocaSenha = null; await carregarSessao(); },
  async pedirAcesso(empresa){
    const {error} = await sb.from("perfis").update({empresa}).eq("id", S.me.id);
    return !error;
  },
  async perfis(ids){
    const o = {}; for(const id of ids){ const p = S.perfis.find(x=>x.id===id); if(p) o[id] = perfilVisivel(p); }
    return o;
  }
};

/* ---------- inicialização ---------- */
async function boot(){
  initForm(); renderMix(); render();
  await DATA.iniciar();
}
boot();
})();
