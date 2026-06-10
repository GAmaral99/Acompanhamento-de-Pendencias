let filialAtual=null, depAtual=null;
let filtroCoord=null, filtroResp=null, filtroGrupo=null;
let filtroVencidas=false;
let abaAtual='man', bfPeriodoAtual='dia';

// ── Labels ────────────────────────────────────────────────────────────────────
const LABELS = {'SP':'São Paulo','Santos':'Santos','RJ':'Rio de Janeiro','GOIAS':'Goiás'};
function labelFilial(f){ return LABELS[f]||f; }

// ── Data helpers ──────────────────────────────────────────────────────────────
function getFiliais(){
  const ocultas = [];
  return [...new Set([...DATA.man,...DATA.fin,...DATA.add].map(r=>r.unidade))]
    .filter(f => !ocultas.includes(f))
    .sort();
}
function getDeps(filial){
  const ocultos = [];
  return [...new Set([...DATA.man,...DATA.fin,...DATA.add]
    .filter(r=>r.unidade===filial)
    .map(r=>r.dep))]
    .filter(d => !ocultos.includes(d.toUpperCase()))
    .sort();
}
function countIn(arr,filial,dep){
  return arr.filter(r=>r.unidade===filial&&r.dep===dep).length;
}

function rowPassaCoord(r){
  if(!filtroCoord) return true;
  if(!r.coord) return false;
  return r.coord.split('|').map(s=>s.trim()).includes(filtroCoord);
}

function parseDateBR(str){
  if(!str) return null;
  const p=str.trim().split('/');
  if(p.length!==3) return null;
  return new Date(+p[2],+p[1]-1,+p[0]);
}

function filtrar(arr, aplicarVencidas=false){
  const hoje=new Date(); hoje.setHours(0,0,0,0);
  return arr.filter(r=>{
    if(r.unidade!==filialAtual||r.dep!==depAtual) return false;
    if(!rowPassaCoord(r)) return false;
    if(filtroResp  && r.resp !==filtroResp)  return false;
    if(filtroGrupo && r.grupo!==filtroGrupo) return false;
    if(aplicarVencidas && filtroVencidas){
      const d=parseDateBR(r.venc);
      if(!d||d>=hoje) return false;
    }
    return true;
  });
}

// ── Screens ───────────────────────────────────────────────────────────────────
function showScreen(id){document.querySelectorAll('.screen').forEach(s=>s.classList.remove('active'));document.getElementById(id).classList.add('active');window.scrollTo(0,0);}
function voltarFilial(){showScreen('screen-filial');filialAtual=null;depAtual=null;}
function voltarDep(){showScreen('screen-dep');depAtual=null;filtroCoord=null;filtroResp=null;filtroGrupo=null;}

// ── Screen 1 ──────────────────────────────────────────────────────────────────
function buildFiliais(){
  const grid=document.getElementById('filial-grid');
  getFiliais().forEach((f,i)=>{
    const total=DATA.man.filter(r=>r.unidade===f).length;
    const c=document.createElement('div');
    c.className='filial-card';
    c.innerHTML=`<div class="fc-num">FILIAL ${String(i+1).padStart(2,'0')}</div><div class="fc-name">${labelFilial(f)}</div><div class="fc-badge">${total} pendência${total!==1?'s':''}</div>`;
    c.onclick=()=>selecionarFilial(f);
    grid.appendChild(c);
  });
}

// ── Screen 2 ──────────────────────────────────────────────────────────────────
function selecionarFilial(filial){
  filialAtual=filial; filtroCoord=null; filtroResp=null; filtroGrupo=null;
  document.getElementById('dep-screen-title').textContent=labelFilial(filial);
  const deps=getDeps(filial);
  document.getElementById('dep-screen-sub').textContent=`${deps.length} departamentos`;
  const list=document.getElementById('dep-list');
  list.innerHTML='';
  deps.forEach(dep=>{
    const m=countIn(DATA.man,filial,dep);
    const f=countIn(DATA.fin,filial,dep);
    const a=countIn(DATA.add,filial,dep);
    const el=document.createElement('div');
    el.className='dep-item';
    el.innerHTML=`<span class="dep-item-name">${dep}</span><div class="dep-pills"><span class="dp dp-man">${m}</span>${MODO_COMP?`<span class="dp dp-fin">${f}</span><span class="dp dp-add">${a}</span>`:''}</div>`;    el.onclick=()=>selecionarDep(dep);
    list.appendChild(el);
  });
  showScreen('screen-dep');
}

// ── Screen 3 ──────────────────────────────────────────────────────────────────
function selecionarDep(dep){
  depAtual=dep; filtroCoord=null; filtroResp=null; filtroGrupo=null;
  abaAtual='man'; bfPeriodoAtual='dia';
  document.getElementById('rh-title').innerHTML=`<em>${labelFilial(filialAtual)}</em> · ${dep}`;
  document.getElementById('rh-meta-base').textContent=`Base: ${ARQ_BASE}`;
  document.getElementById('rh-meta-atual').textContent=MODO_COMP?`Atual: ${ARQ_ATUAL} · Gerado em: ${GERADO}`:`Gerado em: ${GERADO}`;
  document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('ativo'));
  document.getElementById('tbtn-man').classList.add('ativo');
  document.querySelectorAll('.tab-content').forEach(t=>t.classList.remove('ativo'));
  document.getElementById('tab-man').classList.add('ativo');
  document.querySelectorAll('.bf-tab').forEach(b=>b.classList.remove('ativo'));
  document.getElementById('bf-tab-dia').classList.add('ativo');

  buildAllDropdowns();
  renderAll();
  renderPlacares();
  renderBaixasFunc();
  showScreen('screen-results');
}

// ══════════════════════════════════════════════════════════════════════════════
// ── DROPDOWN SYSTEM ───────────────────────────────────────────────────────────
// ══════════════════════════════════════════════════════════════════════════════

function getOpcoesFiltro(tipo){
  const all = [...DATA.man,...DATA.fin,...DATA.add].filter(r => r.unidade===filialAtual && r.dep===depAtual);

  if(tipo==='coord'){
    const cnt={};
    all.forEach(r=>{
      const coords = r.coord ? r.coord.split('|').map(s=>s.trim()).filter(Boolean) : [];
      if(!coords.length){ cnt['']=(cnt['']||0)+1; return; }
      coords.forEach(c=>{ cnt[c]=(cnt[c]||0)+1; });
    });
    return Object.entries(cnt)
      .filter(([k])=>k)
      .sort((a,b)=>b[1]-a[1])
      .map(([k,v])=>{return{valor:k,label:k,count:v}});
  }

  if(tipo==='resp'){
    const filtrados = filtroCoord
      ? all.filter(r=>r.coord && r.coord.split('|').map(s=>s.trim()).includes(filtroCoord))
      : all;
    const cnt={};
    filtrados.forEach(r=>{ if(r.resp) cnt[r.resp]=(cnt[r.resp]||0)+1; });
    return Object.entries(cnt)
      .sort((a,b)=>b[1]-a[1])
      .map(([k,v])=>{return{valor:k,label:k,count:v}});
  }

  if(tipo==='grupo'){
    const filtrados = all.filter(r=>
      rowPassaCoord(r) &&
      (!filtroResp || r.resp===filtroResp)
    );
    const cnt={};
    filtrados.forEach(r=>{ if(r.grupo) cnt[r.grupo]=(cnt[r.grupo]||0)+1; });
    return Object.entries(cnt)
      .sort((a,b)=>b[1]-a[1])
      .map(([k,v])=>{return{valor:k,label:k,count:v}});
  }
  return [];
}

function buildAllDropdowns(){
  renderDropdownOpcoes('coord','');
  renderDropdownOpcoes('resp','');
  renderDropdownOpcoes('grupo','');
  atualizarBotaoFiltro('coord', filtroCoord, 'Todos os coordenadores');
  atualizarBotaoFiltro('resp',  filtroResp,  'Todos os responsáveis');
  atualizarBotaoFiltro('grupo', filtroGrupo, 'Todos os grupos');
}

function renderDropdownOpcoes(tipo, busca){
  const el=document.getElementById(`${tipo}-options`);
  const opcoes=getOpcoesFiltro(tipo);
  const valorAtual = tipo==='coord'?filtroCoord : tipo==='resp'?filtroResp : filtroGrupo;
  const filtradas = busca ? opcoes.filter(o=>o.label.toLowerCase().includes(busca.toLowerCase())) : opcoes;
  if(!filtradas.length){
    el.innerHTML='<div class="filter-opt" style="color:var(--muted);justify-content:center;cursor:default">Nenhuma opção</div>';
    return;
  }
  el.innerHTML='';
  filtradas.forEach(o=>{
    const opt=document.createElement('div');
    opt.className='filter-opt'+(valorAtual===o.valor?' sel':'');
    opt.innerHTML=`<span class="filter-opt-name">${o.label}</span><span class="filter-opt-count">${o.count}</span>`;
    opt.onclick=()=>selecionarFiltro(tipo,o.valor,o.label);
    el.appendChild(opt);
  });
}

function filtrarOpcoes(tipo){
  const busca=document.getElementById(`${tipo}-search-input`).value;
  renderDropdownOpcoes(tipo,busca);
}

function toggleFilter(tipo){
  const btn=document.getElementById(`${tipo}-btn`);
  const dd=document.getElementById(`${tipo}-dropdown`);
  const open=dd.classList.toggle('open');
  btn.classList.toggle('open',open);
  if(open){
    document.getElementById(`${tipo}-search-input`).value='';
    renderDropdownOpcoes(tipo,'');
    setTimeout(()=>document.getElementById(`${tipo}-search-input`).focus(),50);
  }
}

function toggleVencidas(){
  filtroVencidas=!filtroVencidas;
  const btn=document.getElementById('btn-vencidas');
  btn.style.background    = filtroVencidas ? 'rgba(227,30,36,.15)' : 'var(--surface)';
  btn.style.borderColor   = filtroVencidas ? 'var(--red)'          : 'var(--border)';
  btn.style.color         = filtroVencidas ? 'var(--red)'          : 'var(--muted)';
  renderAll(); renderPlacares(); renderBaixasFunc();
}

function fecharTodosDropdowns(){
  ['coord','resp','grupo'].forEach(t=>{
    document.getElementById(`${t}-dropdown`).classList.remove('open');
    document.getElementById(`${t}-btn`).classList.remove('open');
  });
}

function atualizarBotaoFiltro(tipo, valor, placeholder){
  document.getElementById(`${tipo}-btn-label`).textContent = valor || placeholder;
}

function selecionarFiltro(tipo, valor, label){
  if(tipo==='coord'){
    filtroCoord=valor;
    filtroResp=null;
    atualizarBotaoFiltro('coord', valor, 'Todos os coordenadores');
    atualizarBotaoFiltro('resp', null, 'Todos os responsáveis');
    renderDropdownOpcoes('resp','');
    renderDropdownOpcoes('grupo','');
  } else if(tipo==='resp'){
    filtroResp=valor;
    atualizarBotaoFiltro('resp', valor, 'Todos os responsáveis');
    renderDropdownOpcoes('grupo','');
  } else {
    filtroGrupo=valor;
    atualizarBotaoFiltro('grupo', valor, 'Todos os grupos');
  }
  fecharTodosDropdowns();
  renderAll(); renderPlacares(); renderBaixasFunc();
}

function limparFiltro(tipo){
  if(tipo==='coord'){
    filtroCoord=null; filtroResp=null; filtroGrupo=null;
    atualizarBotaoFiltro('coord', null, 'Todos os coordenadores');
    atualizarBotaoFiltro('resp',  null, 'Todos os responsáveis');
    atualizarBotaoFiltro('grupo', null, 'Todos os grupos');
    renderDropdownOpcoes('resp','');
    renderDropdownOpcoes('grupo','');
  } else if(tipo==='resp'){
    filtroResp=null; filtroGrupo=null;
    atualizarBotaoFiltro('resp',  null, 'Todos os responsáveis');
    atualizarBotaoFiltro('grupo', null, 'Todos os grupos');
    renderDropdownOpcoes('grupo','');
  } else {
    filtroGrupo=null;
    atualizarBotaoFiltro('grupo', null, 'Todos os grupos');
  }
  fecharTodosDropdowns();
  renderAll(); renderPlacares(); renderBaixasFunc();
}

document.addEventListener('click',e=>{
  const wraps=['coord-wrap','resp-wrap','grupo-wrap'];
  const inside=wraps.some(id=>{const el=document.getElementById(id);return el&&el.contains(e.target);});
  if(!inside) fecharTodosDropdowns();
});

// ── Render Tables ─────────────────────────────────────────────────────────────
function renderAll(){
  const m=filtrar(DATA.man, true), f=filtrar(DATA.fin), a=filtrar(DATA.add);
  ['man','fin','add'].forEach((k,i)=>{
    const n=[m,f,a][i].length;
    document.getElementById('sc-'+k).textContent=n;
    document.getElementById('bdg-'+k).textContent=n;
  });
  renderTabela('man',m); renderTabela('fin',f); renderTabela('add',a);
}

function renderTabela(aba,rows){
  const tbody=document.getElementById('tbody-'+aba);
  if(!rows.length){tbody.innerHTML='<tr><td colspan="6" class="no-data">Nenhuma tarefa encontrada.</td></tr>';return;}
  const hoje=new Date(); hoje.setHours(0,0,0,0);
  const sorted=[...rows].sort((a,b)=>{
    const da=parseDateBR(a.venc), db=parseDateBR(b.venc);
    if(!da&&!db) return 0;
    if(!da) return 1;
    if(!db) return -1;
    return da-db;
});
const grupos={};
sorted.forEach(r=>{
    const key=r.dep+(r.novo?' (Cliente Novo)':'');
    if(!grupos[key])grupos[key]={dep:r.dep,novo:r.novo,rows:[]};
    grupos[key].rows.push(r);
  });
  let html='';
  Object.keys(grupos).sort().forEach(gkey=>{
    const g=grupos[gkey];
    const nb=g.novo?'<span class="cn-badge">CLIENTE NOVO</span>':'';
    html+=`<tr class="dep-row"><td colspan="6">${g.dep}${nb}</td></tr>`;
    g.rows.forEach(r=>{
      const razaoShort=(r.razao||r.grupo||'—').substring(0,22)+(((r.razao||'').length>22)?'…':'');
      const clientTag=`<span class="client-tag" title="${r.razao||''}"><span class="razao">${razaoShort}</span><span class="cod">#${r.cod}</span></span>`;
      let comtHtml='';
      if(r.dataComt)comtHtml+=`<div class="comt-date">📅 ${r.dataComt}</div>`;
      if(r.comt&&r.comt.trim()){
        comtHtml+=`<button class="comt-btn" onclick="toggleComt(this)">Ver comentário</button>`;
        comtHtml+=`<div class="comt-full">${r.comt.replace(/</g,'&lt;').replace(/>/g,'&gt;')}</div>`;
      }else if(!r.dataComt)comtHtml='<span style="color:var(--muted)">—</span>';
      html+=`<tr><td>${clientTag}</td><td>${r.tit}</td><td style="white-space:nowrap">${r.venc}</td><td class="resp-name">${r.resp||'—'}</td><td style="white-space:nowrap">${r.prev}</td><td>${comtHtml}</td></tr>`;
    });
  });
  tbody.innerHTML=html;
}

function mudarTab(aba,btn){
  document.querySelectorAll('.tab-content').forEach(t=>t.classList.remove('ativo'));
  document.querySelectorAll('.tab-btn').forEach(b=>b.classList.remove('ativo'));
  document.getElementById('tab-'+aba).classList.add('ativo');
  btn.classList.add('ativo'); abaAtual=aba;
}
function toggleComt(btn){
  const open=btn.nextElementSibling.classList.toggle('open');
  btn.textContent=open?'Ocultar':'Ver comentário';
}

// ── Placares de Baixas ────────────────────────────────────────────────────────
function renderPlacares(){
  const sec=document.getElementById('placares-section');
  if(!PLACARES||Object.keys(PLACARES).length===0){sec.style.display='none';return;}
  const temDados=['dia','sem','mes'].some(p=>{
    const pd=PLACARES[p]||{por_resp:[]};
    return pd.por_resp.some(r=>r.unidade===filialAtual&&r.dep===depAtual);
  });
  sec.style.display=temDados?'block':'none';
  ['dia','sem','mes'].forEach(p=>{
    const pd=PLACARES[p]||{total:0,por_resp:[]};
    const filtrados=pd.por_resp.filter(r=>
      r.unidade===filialAtual&&
      r.dep===depAtual&&
      (!filtroCoord||(r.coord||'').split('|').map(s=>s.trim()).includes(filtroCoord))&&
      (!filtroResp||r.resp===filtroResp)
    );
    const total=filtrados.reduce((a,b)=>a+b.n,0);
    document.getElementById(`pc-${p}-num`).textContent=total;
    const container=document.getElementById(`pc-${p}-deps`);
    if(!filtrados.length){container.innerHTML='<div class="pc-empty">Nenhuma baixa neste período</div>';return;}
    const max_n=Math.max(...filtrados.map(r=>r.n),1);
    container.innerHTML=filtrados.map(r=>{
      const pct=Math.round(r.n/max_n*100);
      return `<div class="pc-dep-row"><span class="pc-dep-name" title="${r.resp}">${r.resp}</span><span class="pc-dep-n">${r.n}</span></div><div class="pc-bar-wrap"><div class="pc-bar" style="width:${pct}%"></div></div>`;
    }).join('');
  });
}

// ── Baixas por Funcionário ─────────────────────────────────────────────────────
function renderBaixasFunc(){
  const sec=document.getElementById('baixas-func-section');
  if(!PLACARES||Object.keys(PLACARES).length===0){sec.style.display='none';return;}

  sec.style.display = 'block';

  const pd=PLACARES[bfPeriodoAtual]||{total:0,por_resp:[]};
  const rows=pd.por_resp.filter(r=>
    r.unidade===filialAtual &&
    r.dep===depAtual &&
    (!filtroCoord || (r.coord||'').split('|').map(s=>s.trim()).includes(filtroCoord)) &&
    (!filtroResp  || r.resp===filtroResp)
  );

  // Se o período atual não tem dados, mostra mensagem mas mantém seção visível
  if(!rows.length){
    document.getElementById('bf-tbody').innerHTML=
      '<tr><td colspan="5" class="bf-empty">Nenhuma baixa registrada neste período.</td></tr>';
    return;
  }

  const maxN=Math.max(...rows.map(r=>r.n),1);
  const porDep={};
  rows.forEach(r=>{
    if(!porDep[r.dep]) porDep[r.dep]=[];
    porDep[r.dep].push(r);
  });

  let html='';
  Object.keys(porDep).sort().forEach(dep=>{
    html+=`<tr class="bf-dep-row"><td colspan="5">${dep}</td></tr>`;
    porDep[dep].forEach(r=>{
      const pct=Math.round(r.n/maxN*100);
      const coordLabel=r.coord?r.coord.split('|').map(s=>s.trim()).join(', '):'—';
      html+=`<tr>
        <td><span class="bf-resp-name">${r.resp||'—'}</span></td>
        <td><span class="bf-coord-badge">${coordLabel}</span></td>
        <td style="font-size:12px;color:var(--muted)">${r.dep}</td>
        <td class="bf-bar-cell"><div class="bf-bar-bg"><div class="bf-bar-fill" style="width:${pct}%"></div></div></td>
        <td><span class="bf-num">${r.n}</span></td>
      </tr>`;
    });
  });
  document.getElementById('bf-tbody').innerHTML=html;
}

function mudarBfTab(periodo,btn){
  bfPeriodoAtual=periodo;
  document.querySelectorAll('.bf-tab').forEach(b=>b.classList.remove('ativo'));
  btn.classList.add('ativo');
  renderBaixasFunc();
}

// ── Export Placares ───────────────────────────────────────────────────────────
async function exportarPlacares(modo){
  const overlay=document.createElement('div');
  overlay.className='capture-overlay';
  overlay.innerHTML='<div class="capture-spinner"></div><span style="color:#aaa;font-size:13px">Gerando imagem...</span>';
  document.body.appendChild(overlay);
  await new Promise(r=>setTimeout(r,100));
  try{
    if(modo==='geral'){
      const el=document.getElementById('placares-grid');
      const canvas=await html2canvas(el,{backgroundColor:'#111111',scale:2,useCORS:true,logging:false});
      _downloadCanvas(canvas,`placares_geral_${depAtual||'todos'}.png`);
    } else {
      const periodos=[{id:'pc-dia',label:'Baixas no Dia'},{id:'pc-sem',label:'Baixas na Semana'},{id:'pc-mes',label:'Baixas no Mês'}];
      const capturas=await Promise.all(periodos.map(p=>html2canvas(document.getElementById(p.id),{backgroundColor:'#111111',scale:2,useCORS:true,logging:false})));
      const pad=24,headerH=60;
      const totalW=capturas.reduce((a,c)=>a+c.width,0)+pad*(capturas.length+1);
      const maxH=Math.max(...capturas.map(c=>c.height));
      const tmpCanvas=document.createElement('canvas');
      tmpCanvas.width=totalW; tmpCanvas.height=maxH+pad*2+headerH;
      const ctx=tmpCanvas.getContext('2d');
      ctx.fillStyle='#0a0a0a'; ctx.fillRect(0,0,tmpCanvas.width,tmpCanvas.height);
      ctx.fillStyle='#ffffff'; ctx.font='bold 28px Inter,sans-serif';
      ctx.fillText('MG Contécnica · Placar de Baixas',pad,38);
      ctx.fillStyle='#666'; ctx.font='20px Inter,sans-serif';
      ctx.fillText(`${labelFilial(filialAtual)} · ${depAtual} · ${GERADO}`,pad,62);
      let x=pad;
      capturas.forEach(c=>{ctx.drawImage(c,x,headerH+pad);x+=c.width+pad;});
      _downloadCanvas(tmpCanvas,`placares_dep_${depAtual||'todos'}.png`);
    }
  }catch(e){console.error('Erro ao exportar:',e);alert('Erro ao gerar imagem.');}
  finally{document.body.removeChild(overlay);}
}
function _downloadCanvas(canvas,filename){
  const link=document.createElement('a');link.download=filename;link.href=canvas.toDataURL('image/png');link.click();
}

// ── Tema claro/escuro ─────────────────────────────────────────────────────────
function toggleTema(){
  const claro = document.body.classList.toggle('light');
  localStorage.setItem('tema','relatorio-pendencias');
  localStorage.setItem('tema-claro', claro ? '1' : '0');
  document.getElementById('theme-toggle').textContent = claro ? '☀️' : '🌙';
}
(function(){
  const claro = localStorage.getItem('tema-claro') === '1';
  if(claro) document.body.classList.add('light');
  document.getElementById('theme-toggle').textContent = claro ? '☀️' : '🌙';
})();

// ── PDF HISTÓRICO ─────────────────────────────────────────────────────────────

function abrirModalPDF(){
  document.getElementById('pdf-modal-overlay').classList.add('open');
}
function fecharModalPDF(e){
  if(!e || e.target===document.getElementById('pdf-modal-overlay'))
    document.getElementById('pdf-modal-overlay').classList.remove('open');
}

async function gerarPDF(periodo){
  fecharModalPDF();
  if(!HIST_B64){
    alert('Histórico não disponível. Execute a Action novamente.');
    return;
  }

  const overlay=document.createElement('div');
  overlay.className='capture-overlay';
  overlay.innerHTML='<div class="capture-spinner"></div><span style="color:#aaa;font-size:13px">Gerando PDF...</span>';
  document.body.appendChild(overlay);

  try{
    await _loadScript('https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js');
    await _loadScript('https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.8.2/jspdf.plugin.autotable.min.js');
    await _loadScript('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js');

    const rows = _lerHistorico();
    if(!rows.length){ alert('Nenhum dado encontrado no histórico.'); return; }

    const hoje = new Date();
    const semanaAtual = _semanaISO(hoje);
    const mesAtual = hoje.toISOString().slice(0,7);

    if(periodo === 'mes_semanas'){
      const filtrados = rows.filter(r => r.Mes === mesAtual && r.Filial === filialAtual && r.Departamento === depAtual);
      if(!filtrados.length){
        alert(`Nenhuma baixa registrada para ${labelFilial(filialAtual)} · ${depAtual} neste mês.`);
        return;
      }
      _construirPDFMesSemanas(filtrados, hoje);
      return;
    }

    const filtrados = rows.filter(r => {
      if(periodo==='sem') return r.Semana === semanaAtual && r.Filial === filialAtual && r.Departamento === depAtual;
      return r.Mes === mesAtual && r.Filial === filialAtual && r.Departamento === depAtual;
    });

    if(!filtrados.length){
      alert(`Nenhuma baixa registrada para ${labelFilial(filialAtual)} · ${depAtual} neste período.`);
      return;
    }

    await _construirPDF(filtrados, periodo, hoje);
  } catch(e){
    console.error(e);
    alert('Erro ao gerar PDF: ' + e.message);
  } finally{
    document.body.removeChild(overlay);
  }
}

function _lerHistorico(){
  try{
    const bin = atob(HIST_B64);
    const arr = new Uint8Array(bin.length);
    for(let i=0;i<bin.length;i++) arr[i]=bin.charCodeAt(i);
    const wb = XLSX.read(arr, {type:'array'});
    const ws = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(ws);
  } catch(e){ console.error('Erro ao ler histórico:',e); return []; }
}

function _semanaISO(d){
  const tmp = new Date(Date.UTC(d.getFullYear(), d.getMonth(), d.getDate()));
  const dow = tmp.getUTCDay()||7;
  tmp.setUTCDate(tmp.getUTCDate()+4-dow);
  const y = tmp.getUTCFullYear();
  const w = Math.ceil(((tmp-new Date(Date.UTC(y,0,1)))/86400000+1)/7);
  return `${y}-W${String(w).padStart(2,'0')}`;
}

function _loadScript(src){
  return new Promise((res,rej)=>{
    if(document.querySelector(`script[src="${src}"]`)){res();return;}
    const s=document.createElement('script');
    s.src=src; s.onload=res; s.onerror=rej;
    document.head.appendChild(s);
  });
}

async function _construirPDF(rows, periodo, hoje){
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({orientation:'landscape',unit:'mm',format:'a4'});
  const periodoLabel = periodo==='sem' ? 'Semana Atual' : 'Mês Atual';
  const dataLabel = hoje.toLocaleDateString('pt-BR');

  _cabecalhoPDF(doc, periodoLabel, dataLabel);

  const totais = {};
  const hojeStr = hoje.toISOString().slice(0,10);
  let head, body;
  let funcs = [];

  if(periodo === 'mes'){
    // weekData[resp][sem] = {maxBS: max(BaixasSemana), coord}
    // Para semanas concluídas, max(BaixasSemana) = total da semana (acumula até o último dia).
    // Para a semana atual, max(BaixasSemana) = acumulado até hoje.
    // Propriedade: sum(maxBS por semana) == BaixasMes => soma das colunas == Total Mês.
    const weekData = {};
    rows.forEach(r=>{
      const d = r.Data||'';
      if(!d) return;
      const sem = _semanaDoMes(d);
      const resp = r.Funcionario||'—';
      if(!weekData[resp]) weekData[resp]={};
      if(!weekData[resp][sem]) weekData[resp][sem]={maxBS:0, coord:r.Coordenador||'—'};
      weekData[resp][sem].maxBS = Math.max(weekData[resp][sem].maxBS, Number(r.BaixasSemana)||0);
      if(!totais[resp]) totais[resp]={coord:r.Coordenador||'—', mes:0};
      totais[resp].mes = Math.max(totais[resp].mes, Number(r.BaixasMes)||0);
    });
    // Injeta semana atual via PLACARES.sem (já acumula o dia de hoje)
    const semHoje = _semanaDoMes(hojeStr);
    ((PLACARES.sem||{}).por_resp||[]).forEach(r=>{
      if(r.unidade!==filialAtual||r.dep!==depAtual) return;
      const resp=r.resp||'—';
      if(!weekData[resp]) weekData[resp]={};
      if(!weekData[resp][semHoje]) weekData[resp][semHoje]={maxBS:0, coord:r.coord||'—'};
      weekData[resp][semHoje].maxBS = Math.max(weekData[resp][semHoje].maxBS, r.n||0);
      if(!totais[resp]) totais[resp]={coord:r.coord||'—', mes:0};
    });
    ((PLACARES.mes||{}).por_resp||[]).forEach(r=>{
      if(r.unidade!==filialAtual||r.dep!==depAtual) return;
      const resp=r.resp||'—';
      if(!totais[resp]) totais[resp]={coord:r.coord||'—', mes:0};
      totais[resp].mes = Math.max(totais[resp].mes, r.n||0);
    });

    const allSems=[...new Set(
      Object.values(weekData).flatMap(d=>Object.keys(d).map(Number))
    )].sort((a,b)=>a-b);
    funcs = Object.keys(totais).sort();

    // Coluna "01/MM": BaixasMes - soma de todas as semanas visíveis.
    // Captura baixas anteriores ao primeiro registro do histórico deste mês.
    const labelAnterior = `01/${String(hoje.getMonth()+1).padStart(2,'0')}`;
    const anterior = {};
    funcs.forEach(f=>{
      const somaSemsF = allSems.reduce((a,s)=>a+((weekData[f]?.[s]?.maxBS)||0),0);
      const v = (totais[f].mes||0) - somaSemsF;
      anterior[f] = v > 0 ? v : 0;
    });
    const totalAnterior = funcs.reduce((a,f)=>a+(anterior[f]||0),0);

    head = [['Funcionário','Coordenador',labelAnterior,...allSems.map(s=>`Semana ${s}`),'Total Mês']];
    body = funcs.map(f=>{
      const t=totais[f];
      const wd=weekData[f]||{};
      const antVal = anterior[f]>0 ? String(anterior[f]) : '—';
      const vals=allSems.map(s=>{
        const v=wd[s]?.maxBS;
        return (v!=null&&v>0)?String(v):'—';
      });
      return [f, t.coord, antVal, ...vals, String(t.mes)];
    });
    const totalRow=['TOTAL','', totalAnterior>0?String(totalAnterior):'—'];
    allSems.forEach(s=>totalRow.push(String(funcs.reduce((a,f)=>a+((weekData[f]?.[s]?.maxBS)||0),0))));
    totalRow.push(String(funcs.reduce((a,f)=>a+(totais[f].mes||0),0)));
    body.push(totalRow);

  }else{
    // periodo === 'sem': dias individuais
    const porDia = {};
    rows.forEach(r=>{
      const d = r.Data||'';
      if(!porDia[d]) porDia[d]={};
      const resp = r.Funcionario||'—';
      if(!porDia[d][resp]) porDia[d][resp]={coord:r.Coordenador||'—', n:0};
      porDia[d][resp].n = Math.max(porDia[d][resp].n, Number(r.BaixasDia)||0);
      if(!totais[resp]) totais[resp]={coord:r.Coordenador||'—', sem:0, mes:0};
    });
    // Injeta dados de HOJE
    if(!porDia[hojeStr]) porDia[hojeStr]={};
    ((PLACARES.dia||{}).por_resp||[]).forEach(r=>{
      if(r.unidade!==filialAtual||r.dep!==depAtual) return;
      const resp=r.resp||'—';
      if(!porDia[hojeStr][resp]) porDia[hojeStr][resp]={coord:r.coord||'—', n:0};
      porDia[hojeStr][resp].n = Math.max(porDia[hojeStr][resp].n, r.n||0);
      if(!totais[resp]) totais[resp]={coord:r.coord||'—', sem:0, mes:0};
    });
    // Totais via placares e histórico
    ((PLACARES.sem||{}).por_resp||[]).forEach(r=>{
      if(r.unidade!==filialAtual||r.dep!==depAtual) return;
      const resp=r.resp||'—';
      if(!totais[resp]) totais[resp]={coord:r.coord||'—', sem:0, mes:0};
      totais[resp].sem = r.n||0;
    });
    ((PLACARES.mes||{}).por_resp||[]).forEach(r=>{
      if(r.unidade!==filialAtual||r.dep!==depAtual) return;
      const resp=r.resp||'—';
      if(!totais[resp]) totais[resp]={coord:r.coord||'—', sem:0, mes:0};
      totais[resp].mes = r.n||0;
    });
    rows.forEach(r=>{
      const resp=r.Funcionario||'—';
      if(!totais[resp]) totais[resp]={coord:r.Coordenador||'—', sem:0, mes:0};
      totais[resp].sem = Math.max(totais[resp].sem, Number(r.BaixasSemana)||0);
      totais[resp].mes = Math.max(totais[resp].mes, Number(r.BaixasMes)||0);
    });

    const dias = Object.keys(porDia).sort();
    const funcs = Object.keys(totais).sort();
    const colsDia = dias.map(d=>{
      const dt=new Date(d+'T12:00:00');
      return {data:d, label:dt.toLocaleDateString('pt-BR',{weekday:'short',day:'2-digit',month:'2-digit'})};
    });

    // Dias da semana (segunda até hoje) sem registro no histórico
    const diasSemana=[];
    const dowHoje = hoje.getDay()||7;
    for(let i=1;i<=dowHoje;i++){
      const dt=new Date(hoje);
      dt.setDate(hoje.getDate()-(dowHoje-i));
      diasSemana.push(dt.toISOString().slice(0,10));
    }
    const diasFaltantes = diasSemana.filter(d=>!dias.includes(d));

    // Se BaixasSemana > soma dos dias visíveis, distribui a diferença
    // no único dia faltante, ou agrupa em "Outros" se houver mais de um.
    let colExtra = null;
    if(diasFaltantes.length>=1){
      const diffPorFunc = {};
      funcs.forEach(f=>{
        const somaDias = colsDia.reduce((a,c)=>a+(porDia[c.data]?.[f]?.n||0),0);
        const diff = (totais[f].sem||0) - somaDias;
        diffPorFunc[f] = diff>0?diff:0;
      });
      const label = diasFaltantes.length===1
        ? new Date(diasFaltantes[0]+'T12:00:00').toLocaleDateString('pt-BR',{weekday:'short',day:'2-digit',month:'2-digit'})
        : 'Outros';
      colExtra = {label, valores:diffPorFunc};
    }

    head = [['Funcionário','Coordenador',...colsDia.map(c=>c.label),...(colExtra?[colExtra.label]:[]),'Total Sem','Total Mês']];
    body = funcs.map(f=>{
      const t=totais[f];
      const diasVals=colsDia.map(c=>{const v=porDia[c.data]?.[f]?.n; return v!=null?String(v):'—';});
      const row=[f, t.coord, ...diasVals];
      if(colExtra) row.push(colExtra.valores[f]>0?String(colExtra.valores[f]):'—');
      row.push(String(t.sem), String(t.mes));
      return row;
    });
    const totalRow=['TOTAL',''];
    colsDia.forEach(c=>totalRow.push(String(funcs.reduce((a,f)=>a+(porDia[c.data]?.[f]?.n||0),0))));
    if(colExtra) totalRow.push(String(funcs.reduce((a,f)=>a+(colExtra.valores[f]||0),0)));
    totalRow.push(String(funcs.reduce((a,f)=>a+(totais[f].sem||0),0)));
    totalRow.push(String(funcs.reduce((a,f)=>a+(totais[f].mes||0),0)));
    body.push(totalRow);
  }

  doc.autoTable({
    head, body,
    startY: 24,
    styles:{fontSize:8,cellPadding:2,halign:'center',textColor:[30,30,30]},
    headStyles:{fillColor:[227,30,36],textColor:255,fontStyle:'bold',fontSize:8},
    columnStyles:{0:{halign:'left',fontStyle:'bold',cellWidth:38},1:{halign:'left',cellWidth:28}},
    alternateRowStyles:{fillColor:[245,245,245]},
    didParseCell: function(data){
      if(data.row.index===body.length-1){
        data.cell.styles.fillColor=[227,30,36];
        data.cell.styles.textColor=255;
        data.cell.styles.fontStyle='bold';
      }
    },
    margin:{left:10,right:10},
  });

  _rodapePDF(doc);

  if(periodo === 'mes' && funcs.length){
    doc.addPage();
    _cabecalhoPDF(doc, 'Baixas por Funcionário — Mês', dataLabel);
    await _loadScript('https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js');
    const cv = document.createElement('canvas');
    cv.width = 1600; cv.height = 750;
    Object.assign(cv.style, {position:'absolute',left:'-9999px',top:'-9999px'});
    document.body.appendChild(cv);
    const sortedFuncs = funcs.slice().sort((a,b)=>(totais[b].mes||0)-(totais[a].mes||0));
    const chObj = new Chart(cv, {
      type:'bar',
      data:{
        labels:sortedFuncs.map(f=>f.length>18?f.slice(0,16)+'…':f),
        datasets:[{
          label:'Baixas no Mês',
          data:sortedFuncs.map(f=>totais[f].mes||0),
          backgroundColor:'#E31E24',
          borderColor:'#c01a1f',
          borderWidth:1,
        }],
      },
      options:{
        responsive:false,animation:false,
        plugins:{
          legend:{display:false},
          title:{display:false},
        },
        scales:{
          y:{beginAtZero:true,ticks:{color:'#333',font:{size:11}},grid:{color:'#eee'}},
          x:{ticks:{color:'#333',font:{size:11},maxRotation:45,minRotation:30}},
        },
      },
    });
    await new Promise(r => setTimeout(r, 300));
    const imgData = cv.toDataURL('image/png');
    chObj.destroy(); document.body.removeChild(cv);
    doc.addImage(imgData,'PNG',10,22,277,163);
  }

  const nomeArq=`relatorio_baixas_${depAtual.replace(/\s+/g,'_')}_${periodo}_${hoje.toISOString().slice(0,10)}.pdf`;
  doc.save(nomeArq);
}

function _cabecalhoPDF(doc, periodoLabel, dataLabel){
  doc.setFillColor(227,30,36);
  doc.rect(0,0,297,18,'F');
  doc.setTextColor(255,255,255);
  doc.setFontSize(13); doc.setFont('helvetica','bold');
  doc.text('MG Contécnica · Relatório de Baixas', 10, 12);
  doc.setFontSize(9); doc.setFont('helvetica','normal');
  doc.text(`${labelFilial(filialAtual)} · ${depAtual} · ${periodoLabel} · Gerado em ${dataLabel}`, 10, 17);
}

function _rodapePDF(doc){
  const pageCount = doc.internal.getNumberOfPages();
  for(let i=1; i<=pageCount; i++){
    doc.setPage(i);
    doc.setFontSize(7); doc.setTextColor(150);
    doc.text(`Página ${i} de ${pageCount}`, 287, 205, {align:'right'});
    doc.text('MG Contécnica · Relatório Confidencial', 10, 205);
  }
}

function _semanaDoMes(dataStr){
  return Math.floor((new Date(dataStr+'T12:00:00').getDate()-1)/7)+1;
}

function _construirPDFMesSemanas(rows, hoje){
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({orientation:'landscape', unit:'mm', format:'a4'});
  const mesLabel = hoje.toLocaleDateString('pt-BR', {month:'long', year:'numeric'});
  const dataLabel = hoje.toLocaleDateString('pt-BR');
  const periodoLabel = `Mês por Semanas · ${mesLabel}`;

  // Injeta dados de hoje vindos dos placares em tempo real
  const hojeStr = hoje.toISOString().slice(0, 10);
  const placarDia = (PLACARES.dia || {}).por_resp || [];
  const extraRows = placarDia
    .filter(r => r.unidade === filialAtual && r.dep === depAtual)
    .map(r => ({Data: hojeStr, Funcionario: r.resp||'—', Coordenador: r.coord||'—', BaixasDia: r.n||0}));

  const allRows = [...rows, ...extraRows];

  // Coleta todos os funcionários
  const funcsSet = new Set();
  allRows.forEach(r => funcsSet.add(r.Funcionario||'—'));
  const funcs = [...funcsSet].sort();

  // Agrupa por semana-do-mês e dia
  const semDays = {1:{}, 2:{}, 3:{}, 4:{}};
  const semFuncs = {1:{}, 2:{}, 3:{}, 4:{}};

  allRows.forEach(r => {
    const d = r.Data||'';
    if(!d) return;
    const sem = _semanaDoMes(d);
    const resp = r.Funcionario||'—';
    const val = Number(r.BaixasDia)||0;
    if(!semDays[sem][d]) semDays[sem][d]={};
    semDays[sem][d][resp] = Math.max(semDays[sem][d][resp]||0, val);
    if(!semFuncs[sem][resp]) semFuncs[sem][resp]={coord: r.Coordenador||'—', total:0};
  });

  // Calcula totais semanais por funcionário
  [1,2,3,4].forEach(sem => {
    funcs.forEach(f => {
      if(!semFuncs[sem][f]) semFuncs[sem][f]={coord:'—', total:0};
      semFuncs[sem][f].total = Object.keys(semDays[sem])
        .reduce((a,d) => a+(semDays[sem][d][f]||0), 0);
    });
  });

  // ── Página 1: tabela geral ────────────────────────────────────────────────
  _cabecalhoPDF(doc, periodoLabel, dataLabel);

  const headGeral = [['Funcionário','Coordenador','Semana 1','Semana 2','Semana 3','Semana 4','Total Mês']];
  const bodyGeral = funcs.map(f => {
    const s1=semFuncs[1][f]?.total||0, s2=semFuncs[2][f]?.total||0,
          s3=semFuncs[3][f]?.total||0, s4=semFuncs[4][f]?.total||0;
    const coord = semFuncs[1][f]?.coord||semFuncs[2][f]?.coord||semFuncs[3][f]?.coord||semFuncs[4][f]?.coord||'—';
    return [f, coord, s1||'—', s2||'—', s3||'—', s4||'—', (s1+s2+s3+s4)||'—'];
  });
  bodyGeral.push(['TOTAL','',
    String(funcs.reduce((a,f)=>a+(semFuncs[1][f]?.total||0),0)),
    String(funcs.reduce((a,f)=>a+(semFuncs[2][f]?.total||0),0)),
    String(funcs.reduce((a,f)=>a+(semFuncs[3][f]?.total||0),0)),
    String(funcs.reduce((a,f)=>a+(semFuncs[4][f]?.total||0),0)),
    String(funcs.reduce((a,f)=>{
      return a+(semFuncs[1][f]?.total||0)+(semFuncs[2][f]?.total||0)
              +(semFuncs[3][f]?.total||0)+(semFuncs[4][f]?.total||0);
    },0)),
  ]);

  doc.autoTable({
    head: headGeral, body: bodyGeral,
    startY: 24,
    styles:{fontSize:9,cellPadding:2.5,halign:'center',textColor:[30,30,30]},
    headStyles:{fillColor:[227,30,36],textColor:255,fontStyle:'bold'},
    columnStyles:{0:{halign:'left',fontStyle:'bold',cellWidth:45},1:{halign:'left',cellWidth:35}},
    alternateRowStyles:{fillColor:[245,245,245]},
    didParseCell: function(data){
      if(data.row.index===bodyGeral.length-1){
        data.cell.styles.fillColor=[227,30,36];
        data.cell.styles.textColor=255;
        data.cell.styles.fontStyle='bold';
      }
    },
    margin:{left:10,right:10},
  });

  // ── Páginas 2-5: uma por semana ───────────────────────────────────────────
  [1,2,3,4].forEach(sem => {
    const days = Object.keys(semDays[sem]).sort();
    if(!days.length) return;

    doc.addPage();
    _cabecalhoPDF(doc, `${periodoLabel} · Semana ${sem}`, dataLabel);

    const colsDia = days.map(d => {
      const dt = new Date(d+'T12:00:00');
      return {data:d, label:dt.toLocaleDateString('pt-BR',{weekday:'short',day:'2-digit',month:'2-digit'})};
    });

    const headSem = [['Funcionário','Coordenador',...colsDia.map(c=>c.label),'Total Semana']];
    const bodySem = funcs.map(f => {
      const dvals = colsDia.map(c => {
        const v = semDays[sem][c.data]?.[f];
        return v ? String(v) : '—';
      });
      const total = colsDia.reduce((a,c)=>a+(semDays[sem][c.data]?.[f]||0),0);
      return [f, semFuncs[sem][f]?.coord||'—', ...dvals, total||'—'];
    });
    const totalSem = ['TOTAL',''];
    colsDia.forEach(c => totalSem.push(String(funcs.reduce((a,f)=>a+(semDays[sem][c.data]?.[f]||0),0))));
    totalSem.push(String(funcs.reduce((a,f)=>a+(semFuncs[sem][f]?.total||0),0)));
    bodySem.push(totalSem);

    doc.autoTable({
      head: headSem, body: bodySem,
      startY: 24,
      styles:{fontSize:8,cellPadding:2,halign:'center',textColor:[30,30,30]},
      headStyles:{fillColor:[227,30,36],textColor:255,fontStyle:'bold',fontSize:8},
      columnStyles:{0:{halign:'left',fontStyle:'bold',cellWidth:38},1:{halign:'left',cellWidth:28}},
      alternateRowStyles:{fillColor:[245,245,245]},
      didParseCell: function(data){
        if(data.row.index===bodySem.length-1){
          data.cell.styles.fillColor=[227,30,36];
          data.cell.styles.textColor=255;
          data.cell.styles.fontStyle='bold';
        }
      },
      margin:{left:10,right:10},
    });
  });

  _rodapePDF(doc);

  const nomeArq=`relatorio_baixas_${depAtual.replace(/\s+/g,'_')}_mes_semanas_${hoje.toISOString().slice(0,10)}.pdf`;
  doc.save(nomeArq);
}

// ── Init ──────────────────────────────────────────────────────────────────────
buildFiliais();
