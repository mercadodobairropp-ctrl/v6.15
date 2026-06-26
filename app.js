const Mem={usuarios:null,turnos:null,empresas:null,checklists:null,execucoes:null,last:{usuarios:0,turnos:0,empresas:0,checklists:0,execucoes:0},pending:{}};
const NET={jsonpTimeout:22000,jsonpRetryDelay:650,postTimeout:18000};
function prepararVersaoLocal(){
  const v=localStorage.getItem("appVersaoDados");
  const versaoDados=APP.DADOS_VERSAO||APP.versao;
  if(v===versaoDados)return;
  const preservar=String(v||"").startsWith("v6.15-base-nova-")?new Set(["empresaDispositivo"]):new Set();
  Object.keys(localStorage).forEach(k=>{
    if(preservar.has(k))return;
    if(k.startsWith("exec_")||k.startsWith("rascunho_")||k.includes("_alerta_v")||["empresaDispositivo","usuariosSistema","turnosSistema","empresasSistema","checklistsSistema","filaEnvio","filaArquivosEnvio","syncRevision","syncServidor","ultimaSync","turnoSelecionado"].includes(k))localStorage.removeItem(k);
  });
  localStorage.setItem("appVersaoDados",versaoDados);
}
prepararVersaoLocal();
function desativarSugestoesCampos(root=document){
  const filhos=root.querySelectorAll?[...root.querySelectorAll("input, textarea")]:[];
  const campos=root.matches&&root.matches("input, textarea")?[root,...filhos]:filhos;
  campos.forEach(campo=>{
    campo.setAttribute("autocomplete",campo.type==="password"?"new-password":"off");
    campo.setAttribute("autocorrect","off");
    campo.setAttribute("autocapitalize","off");
    campo.setAttribute("spellcheck","false");
  });
}
document.addEventListener("DOMContentLoaded",()=>desativarSugestoesCampos());
document.addEventListener("focusin",e=>{if(e.target?.matches?.("input, textarea"))desativarSugestoesCampos(e.target)});
function normalizarLogin(v){return String(v||"").trim()}
function normalizarSenha(v){const s=String(v??"").trim();return /^\d+\.0$/.test(s)?s.replace(/\.0$/,""):s}
function normalizarUsuario(u){
  const login=normalizarLogin(u.login);
  let nome=String(u.nome||u.nomeUsuario||u.usuario||"").trim();
  if(!nome && login===APP.ADMIN_MESTRE) nome="André";
  if(!nome) nome="Usuário";
  return {...u,login,nome,tipo:String(u.tipo||"operador").trim().toLowerCase(),turnosPermitidos:parseLista(u.turnosPermitidos||"").join(","),empresasPermitidas:parseLista(u.empresasPermitidas||"").join(",")};
}
function usuarioLogado(){try{const u=JSON.parse(localStorage.getItem("usuarioLogado"));return u?normalizarUsuario(u):null}catch(e){return null}}
function exigirLogin(){const u=usuarioLogado();if(!u){location.href="login.html";return null}return u}
function nomePorLogin(login){
  login=normalizarLogin(login);
  try{
    const lista=JSON.parse(localStorage.getItem("usuariosSistema")||"[]");
    const u=lista.find(x=>normalizarLogin(x.login)===login);
    if(u && u.nome) return u.nome;
  }catch(e){}
  if(login===APP.ADMIN_MESTRE) return "André";
  return login ? "Usuário " + login : "";
}
function nomeExecucao(ex){
  if(!ex) return "";
  return ex.nomeUsuario || ex.nome || nomePorLogin(ex.login);
}
function hojeISO(d=new Date()){const p=n=>String(n).padStart(2,"0");return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())}`}
function dataBR(d=new Date()){return d.toLocaleDateString("pt-BR")}
function horaBR(d=new Date()){return d.toLocaleTimeString("pt-BR",{hour:"2-digit",minute:"2-digit"})}
function horaArquivo(d=new Date()){return horaBR(d).replace(":","-")}
function horarioParaMinutos(h){const s=String(h||"00:00").slice(0,5);const p=s.split(":");return parseInt(p[0]||"0")*60+parseInt(p[1]||"0")}
function agoraMinutos(){const a=new Date();return a.getHours()*60+a.getMinutes()}
function reaberturaExpirada(ex){
  if(!ex)return false;
  if(ex.reabertoEm){
    const base=new Date(ex.reabertoEm);
    if(!isNaN(base.getTime()))return Date.now()>base.getTime()+60*60*1000;
  }
  const lim=horarioParaMinutos(ex.novoHorarioFim||"");
  const now=agoraMinutos();
  return lim?now>lim:false;
}
function diaSemanaAtual(){return ["dom","seg","ter","qua","qui","sex","sab"][new Date().getDay()]}
function parseLista(v){if(Array.isArray(v))return v.map(x=>String(x).trim()).filter(Boolean);return String(v||"").split(/,|;|\n/).map(x=>x.trim()).filter(Boolean)}
function parseTarefas(v){if(Array.isArray(v))return v.flatMap(x=>String(x).replace(/\\n/g,"\n").split(/\n|;/)).map(x=>x.trim()).filter(Boolean);return String(v||"").replace(/\\n/g,"\n").split(/\n|;/).map(x=>x.trim()).filter(Boolean)}
function normalizarHora(v){
  const s=String(v||"").trim();
  if(/^\d{1,2}$/.test(s)){
    const h=Math.min(23,Math.max(0,parseInt(s,10)));
    return String(h).padStart(2,"0")+":00";
  }
  const m=s.match(/(\d{1,2}):?(\d{2})?/);
  if(m){
    const h=Math.min(23,Math.max(0,parseInt(m[1]||"0",10)));
    const min=Math.min(59,Math.max(0,parseInt(m[2]||"0",10)));
    return String(h).padStart(2,"0")+":"+String(min).padStart(2,"0");
  }
  return "00:00";
}
function checklistCompleto(c){return !!(c&&String(c.id||"").trim()&&String(c.nome||"").trim()&&String(c.empresaId||"").trim()&&parseLista(c.turnos).length&&parseTarefas(c.tarefas).length&&normalizarHora(c.horario)!=="00:00"&&normalizarHora(c.horarioFim)!=="00:00")}
function motivoChecklistIncompleto(c){const m=[];if(!String(c.id||"").trim())m.push("ID");if(!String(c.nome||"").trim())m.push("nome");if(!String(c.empresaId||"").trim())m.push("empresa");if(!parseLista(c.turnos).length)m.push("turno");if(!parseTarefas(c.tarefas).length)m.push("atividades");if(normalizarHora(c.horario)==="00:00")m.push("horário inicial");if(normalizarHora(c.horarioFim)==="00:00")m.push("horário final");return m.join(", ")}
function normalizarChecklist(c){const o={id:String(c.id||"").trim(),nome:String(c.nome||"").trim(),descricao:String(c.descricao||"").trim(),horario:normalizarHora(c.horario),horarioFim:normalizarHora(c.horarioFim||"23:59"),turnos:parseLista(c.turnos),dias:parseLista(String(c.dias||"").toLowerCase()),prioridade:String(c.prioridade||"media").toLowerCase(),tarefas:parseTarefas(c.tarefas),ativo:String(c.ativo||"sim").toLowerCase(),empresaId:String(c.empresaId||"").trim()};if(!checklistCompleto(o))o.ativo="nao";return o}
function normalizarTurno(t){return{id:String(t.id||"").trim(),nome:String(t.nome||"").trim(),empresaId:String(t.empresaId||"").trim(),ativo:String(t.ativo||"sim").toLowerCase()}}
function normalizarEmpresa(e){return{id:String(e.id||"").trim(),nome:String(e.nome||"").trim(),ativo:String(e.ativo||"sim").toLowerCase()}}
function execId(c,data=hojeISO()){return `${data}_${c.id}_${String(c.horario).replace(":","-")}`}
function getExecLocal(id){try{return JSON.parse(localStorage.getItem("exec_"+id)||"null")}catch(e){return null}}
function limparArmazenamentoPesado(){
  const hoje=hojeISO();
  Object.keys(localStorage).forEach(k=>{
    if(k.startsWith("rascunho_"))localStorage.removeItem(k);
    if(k.startsWith("exec_")&&!k.includes(hoje))localStorage.removeItem(k);
    if(k.includes("_alerta_v"))localStorage.removeItem(k);
  });
}
function setJSONLocalSeguro(key,d){
  const valor=JSON.stringify(d);
  try{localStorage.setItem(key,valor);return true}catch(e){}
  limparArmazenamentoPesado();
  try{localStorage.setItem(key,valor);return true}catch(e){return false}
}
function salvarCacheLocal(key,d){try{localStorage.setItem(key,JSON.stringify(d));return true}catch(e){return false}}
function salvarSessaoLocal(key,d){if(setJSONLocalSeguro(key,d))return true;["usuariosSistema","turnosSistema","empresasSistema","checklistsSistema"].forEach(k=>localStorage.removeItem(k));return setJSONLocalSeguro(key,d)}
function setExecLocal(id,d){setJSONLocalSeguro("exec_"+id,d)}
function removeExecLocal(id){localStorage.removeItem("exec_"+id)}
function postAPI(data,timeoutMs=NET.postTimeout,url=APP.API_URL){
  const envio=fetch(url,{method:"POST",mode:"no-cors",body:new URLSearchParams(data)});
  if(!timeoutMs)return envio;
  return Promise.race([envio,new Promise(resolve=>setTimeout(()=>resolve({timeout:true}),timeoutMs))]);
}
function jsonp(acao,params={},timeoutMs=NET.jsonpTimeout){return new Promise((resolve,reject)=>{const cb="cb_"+Date.now()+"_"+Math.floor(Math.random()*99999);const qs=new URLSearchParams({...params,acao,callback:cb,_:Date.now()});const s=document.createElement("script");let done=false;const limpar=()=>{clearTimeout(timer);delete window[cb];s.remove()};const timer=setTimeout(()=>{if(done)return;done=true;limpar();reject(new Error("Tempo esgotado ao conectar ao servidor"))},timeoutMs);window[cb]=d=>{if(done)return;done=true;limpar();resolve(d)};s.onerror=()=>{if(done)return;done=true;limpar();reject(new Error("Falha ao conectar ao servidor"))};s.src=APP.API_URL+"?"+qs.toString();document.body.appendChild(s)})}
async function jsonpRetry(acao,params={},tentativas=2,timeoutMs=NET.jsonpTimeout){let erro;for(let i=0;i<tentativas;i++){try{return await jsonp(acao,params,timeoutMs)}catch(e){erro=e;if(i<tentativas-1)await new Promise(r=>setTimeout(r,NET.jsonpRetryDelay))}}throw erro}
function cacheOk(tipo){return Date.now()-(Mem.last[tipo]||0)<APP.CACHE_MS}
function pendente(chave,fn){if(Mem.pending[chave])return Mem.pending[chave];Mem.pending[chave]=fn().finally(()=>delete Mem.pending[chave]);return Mem.pending[chave]}
async function carregarUsuariosOnline(force=false){if(!force&&Mem.usuarios&&cacheOk("usuarios"))return Mem.usuarios;return pendente("usuarios",async()=>{const d=await jsonp("getUsers");if(d?.status==="ok"){Mem.usuarios=(d.usuarios||[]).map(normalizarUsuario);salvarCacheLocal("usuariosSistema",Mem.usuarios);Mem.last.usuarios=Date.now();return Mem.usuarios}throw new Error("Erro ao carregar usuários")})}
async function carregarTurnosOnline(force=false){if(!force&&Mem.turnos&&cacheOk("turnos"))return Mem.turnos;return pendente("turnos",async()=>{const d=await jsonp("getTurnos");if(d?.status==="ok"){Mem.turnos=(d.turnos||[]).map(normalizarTurno).filter(t=>t.id&&t.ativo!=="nao");salvarCacheLocal("turnosSistema",Mem.turnos);Mem.last.turnos=Date.now();return Mem.turnos}throw new Error("Erro ao carregar turnos")})}
async function carregarEmpresasOnline(force=false){if(!force&&Mem.empresas&&cacheOk("empresas"))return Mem.empresas;return pendente("empresas",async()=>{const d=await jsonp("getEmpresas");if(d?.status==="ok"){Mem.empresas=(d.empresas||[]).map(normalizarEmpresa).filter(e=>e.id&&e.ativo!=="nao");salvarCacheLocal("empresasSistema",Mem.empresas);Mem.last.empresas=Date.now();return Mem.empresas}throw new Error("Erro ao carregar empresas")})}
async function carregarChecklistsOnline(force=false,incluirInativos=false){if(!force&&Mem.checklists&&cacheOk("checklists"))return incluirInativos?Mem.checklists:Mem.checklists.filter(c=>c.ativo!=="nao");const lista=await pendente("checklists",async()=>{const d=await jsonp("getChecklists");if(d?.status==="ok"){Mem.checklists=(d.checklists||[]).map(normalizarChecklist).filter(c=>c.id&&c.nome);salvarCacheLocal("checklistsSistema",Mem.checklists);Mem.last.checklists=Date.now();return Mem.checklists}throw new Error("Erro ao carregar checklists")});return incluirInativos?lista:lista.filter(c=>c.ativo!=="nao")}
async function carregarExecucoesOnline(inicio="",fim="",force=false){const chave=`execucoes_${inicio}_${fim}`;if(!force&&Mem.execucoes&&cacheOk("execucoes"))return Mem.execucoes;return pendente(chave,async()=>{const d=await jsonp("getExecucoes",{inicio,fim});if(d?.status==="ok"){Mem.execucoes=d.execucoes||[];Mem.last.execucoes=Date.now();return Mem.execucoes}return[]})}
async function carregarPacoteOnline(inicio=hojeISO(),fim=hojeISO(),force=false){
  const chave=`base_${inicio}_${fim}`;
  if(!force&&Mem.turnos&&Mem.empresas&&Mem.checklists&&Mem.execucoes&&cacheOk("turnos")&&cacheOk("empresas")&&cacheOk("checklists")&&cacheOk("execucoes"))return{sync:null,turnos:Mem.turnos,empresas:Mem.empresas,checklists:Mem.checklists,execucoes:Mem.execucoes};
  return pendente(chave,async()=>{
    try{
      const d=await jsonpRetry("getBase",{inicio,fim},2,force?26000:16000);
      if(d?.status==="ok"&&Array.isArray(d.turnos)&&Array.isArray(d.empresas)&&Array.isArray(d.checklists)){
        const agora=Date.now();
        if(Array.isArray(d.usuarios)){Mem.usuarios=(d.usuarios||[]).map(normalizarUsuario);salvarCacheLocal("usuariosSistema",Mem.usuarios);Mem.last.usuarios=agora}
        Mem.turnos=(d.turnos||[]).map(normalizarTurno).filter(t=>t.id&&t.ativo!=="nao");
        Mem.empresas=(d.empresas||[]).map(normalizarEmpresa).filter(e=>e.id&&e.ativo!=="nao");
        Mem.checklists=(d.checklists||[]).map(normalizarChecklist).filter(c=>c.id&&c.nome);
        Mem.execucoes=d.execucoes||[];
        salvarCacheBase(d,agora);
        return{sync:d,turnos:Mem.turnos,empresas:Mem.empresas,checklists:Mem.checklists,execucoes:Mem.execucoes};
      }
      throw new Error("Erro ao carregar base");
    }catch(e){throw e}
  });
}
async function carregarOperacionalOnline(login,empresaId,data=hojeISO(),force=false){
  const chave=`operacional_${normalizarLogin(login)}_${empresaId}_${data}`;
  if(!force&&Mem.turnos&&Mem.empresas&&Mem.checklists&&Mem.execucoes&&cacheOk("turnos")&&cacheOk("empresas")&&cacheOk("checklists")&&cacheOk("execucoes"))return{sync:null,turnos:Mem.turnos,empresas:Mem.empresas,checklists:Mem.checklists,execucoes:Mem.execucoes,usuario:(Mem.usuarios||[])[0]||null};
  return pendente(chave,async()=>{
    const d=await jsonpRetry("getOperacional",{login:normalizarLogin(login),empresaId,data},2,force?24000:15000);
    if(d?.status==="ok"||d?.status==="bloqueado"){
      const agora=Date.now();
      const usuario=d.usuario?normalizarUsuario(d.usuario):null;
      Mem.usuarios=usuario?[usuario]:[];
      Mem.empresas=(d.empresa?[d.empresa]:(d.empresas||[])).map(normalizarEmpresa).filter(e=>e.id&&e.ativo!=="nao");
      Mem.turnos=(d.turnos||[]).map(normalizarTurno).filter(t=>t.id&&t.ativo!=="nao");
      Mem.checklists=(d.checklists||[]).map(normalizarChecklist).filter(c=>c.id&&c.nome);
      Mem.execucoes=d.execucoes||[];
      salvarCacheLocal("usuariosSistema",Mem.usuarios);
      salvarCacheBase(d,agora);
      return{sync:d,turnos:Mem.turnos,empresas:Mem.empresas,checklists:Mem.checklists,execucoes:Mem.execucoes,usuario,bloqueio:d.status==="bloqueado"?d:null};
    }
    throw new Error(d?.mensagem||"Erro ao carregar dados operacionais");
  });
}
function limparMemoriaSincronizacao(){
  Mem.usuarios=null;Mem.turnos=null;Mem.empresas=null;Mem.checklists=null;Mem.execucoes=null;Mem.pending={};
  Mem.last={usuarios:0,turnos:0,empresas:0,checklists:0,execucoes:0};
}
async function sincronizarBaseCompleta(inicio=hojeISO(),fim=hojeISO()){
  limparMemoriaSincronizacao();
  const pacote=await carregarPacoteOnline(inicio,fim,true);
  if(pacote.offline)throw new Error("Base online indisponível");
  if(!Mem.usuarios)await carregarUsuariosOnline(true);
  return pacote;
}
async function iniciarExecucaoOnline(payload){
  const r=await jsonpRetry("iniciarExecucao",payload,2,26000);
  if(r?.status==="ok")return r;
  const msg=r?.mensagem||"Não foi possível iniciar este checklist.";
  const erro=new Error(msg);
  erro.resposta=r;
  throw erro;
}
function salvarCacheBase(d,agora=Date.now()){salvarCacheLocal("turnosSistema",Mem.turnos||[]);salvarCacheLocal("empresasSistema",Mem.empresas||[]);salvarCacheLocal("checklistsSistema",Mem.checklists||[]);try{if(d?.revision)localStorage.setItem("syncRevision",d.revision||"");if(d?.servidorEm)localStorage.setItem("syncServidor",d.servidorEm||"")}catch(e){}["turnos","empresas","checklists","execucoes"].forEach(k=>Mem.last[k]=agora)}
function listaLocal(key,def=[]){try{const v=localStorage.getItem(key);return v?JSON.parse(v):def}catch(e){return def}}
function baseLocal(){return{turnos:listaLocal("turnosSistema",[]),empresas:listaLocal("empresasSistema",[]),checklists:listaLocal("checklistsSistema",[]),execucoes:Mem.execucoes||[]}}
function abrirBancoDispositivo(){return new Promise((resolve,reject)=>{if(!("indexedDB"in window)){reject(new Error("IndexedDB indisponível"));return}const req=indexedDB.open("checklistDispositivoV615BaseNova",1);req.onupgradeneeded=()=>{const db=req.result;if(!db.objectStoreNames.contains("config"))db.createObjectStore("config",{keyPath:"key"})};req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error||new Error("Falha ao abrir configuração do dispositivo"))})}
async function solicitarPersistenciaArmazenamento(){try{if(!navigator.storage?.persist)return false;const ja=await navigator.storage.persisted?.();if(ja){localStorage.setItem("armazenamentoPersistente","sim");return true}const ok=await navigator.storage.persist();localStorage.setItem("armazenamentoPersistente",ok?"sim":"nao");return ok}catch(e){return false}}
async function salvarConfigDispositivo(key,value){try{const db=await abrirBancoDispositivo();await new Promise((resolve,reject)=>{const tx=db.transaction("config","readwrite");tx.objectStore("config").put({key,value,atualizadoEm:new Date().toISOString()});tx.oncomplete=resolve;tx.onerror=()=>reject(tx.error)});db.close()}catch(e){}}
async function lerConfigDispositivo(key){try{const db=await abrirBancoDispositivo();const item=await new Promise((resolve,reject)=>{const tx=db.transaction("config","readonly"),req=tx.objectStore("config").get(key);req.onsuccess=()=>resolve(req.result||null);req.onerror=()=>reject(req.error)});db.close();return item?.value||""}catch(e){return""}}
async function salvarEmpresaDispositivoPersistente(id){if(id)localStorage.setItem("empresaDispositivo",id);else localStorage.removeItem("empresaDispositivo");await salvarConfigDispositivo("empresaDispositivo",id||"");solicitarPersistenciaArmazenamento().catch(()=>{});return id||""}
async function garantirEmpresaDispositivo(){let id=localStorage.getItem("empresaDispositivo")||"";if(id)return id;id=await lerConfigDispositivo("empresaDispositivo");if(id)localStorage.setItem("empresaDispositivo",id);return id||""}
garantirEmpresaDispositivo().catch(()=>{});
async function obterChecklists(incluirInativos=false){return await carregarChecklistsOnline(true,incluirInativos)}
function usuarioPodeVerTurno(u,t){if(u.tipo==="admin")return true;const p=parseLista(u.turnosPermitidos||"");return p.includes("todos")||p.includes(t)}
function usuarioPodeVerEmpresa(u,e){if(u.tipo==="admin")return true;if(!e)return false;const p=parseLista(u.empresasPermitidas||"");return p.includes("todos")||p.includes(e)}
function checklistPermitidoUsuario(c,u,t){if(u.tipo==="admin")return true;if(!c.turnos.includes(t))return false;return usuarioPodeVerEmpresa(u,c.empresaId)}
function statusExecucao(c,ex=null){const now=agoraMinutos(),ini=horarioParaMinutos(c.horario),fim=horarioParaMinutos(c.horarioFim);if(ex){if(["finalizado","aguardando_envio"].includes(ex.status))return ex.status;if(ex.status==="executando")return(ex.novoHorarioFim||ex.reabertoEm)&&reaberturaExpirada(ex)?"expirado":"executando";if(ex.status==="reaberto")return reaberturaExpirada(ex)?"expirado":"reaberto";if(ex.status==="expirado")return"expirado"}if(now<ini)return"aguardando";if(now>fim)return"expirado";if(now-ini>=60)return"critico";if(now-ini>=30)return"atrasado";return"liberado"}
function classeStatus(st){return{finalizado:"green",aguardando_envio:"blue",executando:"green",reaberto:"green",aguardando:"gray",liberado:"green",atrasado:"red",critico:"darkred",expirado:"gray"}[st]||"gray"}
function textoStatus(st,c,ex){const nome=nomeExecucao(ex);if(st==="finalizado")return`✅ Finalizado${nome?" por "+nome:""}`;if(st==="aguardando_envio")return`🟡 Aguardando envio${nome?" ("+nome+")":""}`;if(st==="executando")return`🟢 Em execução${nome?" por "+nome:""}`;if(st==="reaberto")return`🟢 Reaberto até ${ex?.novoHorarioFim||c.horarioFim}`;if(st==="aguardando")return"⏳ Aguardando horário";if(st==="liberado")return"🟢 Liberado";if(st==="atrasado")return"🔴 Atrasado";if(st==="critico")return"⚫ Crítico";if(st==="expirado")return"❌ Não feito / expirado";return"Pendente"}
function ordenarTexto(a,b){return String(a||"").localeCompare(String(b||""),"pt-BR",{numeric:true,sensitivity:"base"})}
function ordenarChecklistsBase(a,b){return ordenarTexto(a.empresaId,b.empresaId)||ordenarTexto((a.turnos||[])[0],(b.turnos||[])[0])||horarioParaMinutos(a.horario)-horarioParaMinutos(b.horario)||ordenarTexto(a.nome,b.nome)||ordenarTexto(a.id,b.id)}
function ordenarCards(a,b){const fa=["finalizado","expirado","aguardando_envio"].includes(a.status)?1:0,fb=["finalizado","expirado","aguardando_envio"].includes(b.status)?1:0;if(fa!==fb)return fa-fb;return ordenarChecklistsBase(a.checklist,b.checklist)}
const AlarmesAtivos={timers:[],audios:[]};
function silenciarAlarmeAtual(){AlarmesAtivos.timers.forEach(t=>clearTimeout(t));AlarmesAtivos.timers=[];AlarmesAtivos.audios.forEach(a=>{try{a.pause();a.currentTime=0}catch(e){}});AlarmesAtivos.audios=[];try{navigator.vibrate?.(0)}catch(e){}}
function tocarAlarme(c=1){for(let i=0;i<c;i++){const timer=setTimeout(()=>{try{navigator.vibrate?.(2000)}catch(e){}try{const a=new Audio("https://actions.google.com/sounds/v1/alarms/alarm_clock.ogg");AlarmesAtivos.audios.push(a);a.play();const stop=setTimeout(()=>{try{a.pause();a.currentTime=0}catch(e){}AlarmesAtivos.audios=AlarmesAtivos.audios.filter(x=>x!==a)},2000);AlarmesAtivos.timers.push(stop)}catch(e){}},i*2600);AlarmesAtivos.timers.push(timer)}}
function alertaUnico(id,tipo,ciclos,msg=null){const k=`${id}_${tipo}_alerta_v6`;if(localStorage.getItem(k))return;localStorage.setItem(k,"1");tocarAlarme(ciclos)}
function nomePdf(c){const a=new Date(),ano=a.getFullYear(),mes=String(a.getMonth()+1).padStart(2,"0"),dia=String(a.getDate()).padStart(2,"0");return`${ano}-${mes}-${dia} - ${c.nome.toLowerCase()} - ${horaArquivo(a)}.pdf`}
function marcarSync(t=""){localStorage.setItem("ultimaSync",t||new Date().toISOString())}
function textoUltimaSync(){const s=localStorage.getItem("ultimaSync");if(!s)return"Ainda não sincronizado";const m=Math.floor((Date.now()-new Date(s).getTime())/60000);return m<=0?"Atualizado agora":`Atualizado há ${m} min`}
function proximaMeiaNoiteMs(){const d=new Date();d.setHours(24,0,8,0);return Math.max(1000,d.getTime()-Date.now())}
async function sincronizarDiario(chave,fn){const hoje=hojeISO();if(localStorage.getItem(chave)===hoje)return false;await fn();localStorage.setItem(chave,hoje);return true}
function agendarSyncDiario(chave,fn){sincronizarDiario(chave,fn).catch(()=>{});setTimeout(function disparar(){sincronizarDiario(chave,fn).catch(()=>{});setTimeout(disparar,proximaMeiaNoiteMs())},proximaMeiaNoiteMs())}
function confirmarAcao({titulo="Confirmar ação",texto="",confirmar="Confirmar",cancelar="Cancelar",perigo=false}={}){return new Promise(resolve=>{let modal=document.getElementById("confirmModal");if(!modal){modal=document.createElement("div");modal.id="confirmModal";modal.className="modal";modal.innerHTML=`<div class="modalContent confirmModal"><h2 id="confirmTitulo"></h2><p id="confirmTexto"></p><div class="confirmActions"><button id="confirmCancelar" class="secondary" type="button"></button><button id="confirmOk" type="button"></button></div></div>`;document.body.appendChild(modal)}const ok=document.getElementById("confirmOk"),cancel=document.getElementById("confirmCancelar");document.getElementById("confirmTitulo").innerText=titulo;document.getElementById("confirmTexto").innerText=texto;ok.innerText=confirmar;ok.className=perigo?"red":"";cancel.innerText=cancelar;modal.style.display="flex";const fechar=v=>{modal.style.display="none";ok.onclick=null;cancel.onclick=null;resolve(v)};ok.onclick=()=>fechar(true);cancel.onclick=()=>fechar(false)})}
function abrirBancoUploads(){return new Promise((resolve,reject)=>{if(!("indexedDB"in window)){reject(new Error("IndexedDB indisponível"));return}const req=indexedDB.open("checklistArquivosV615BaseNova",1);req.onupgradeneeded=()=>{const db=req.result;if(!db.objectStoreNames.contains("arquivos"))db.createObjectStore("arquivos",{keyPath:"id"})};req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error||new Error("Falha ao abrir fila de arquivos"))})}
async function idsFilaArquivos(){try{const db=await abrirBancoUploads();const ids=await new Promise((resolve,reject)=>{const tx=db.transaction("arquivos","readonly"),req=tx.objectStore("arquivos").getAllKeys();req.onsuccess=()=>resolve(req.result||[]);req.onerror=()=>reject(req.error)});db.close();return ids}catch(e){return[]}}
async function enfileirarArquivoEnvio(item){const db=await abrirBancoUploads();await new Promise((resolve,reject)=>{const tx=db.transaction("arquivos","readwrite");tx.objectStore("arquivos").put({...item,tentativas:item.tentativas||0,criadoEm:item.criadoEm||new Date().toISOString()});tx.oncomplete=resolve;tx.onerror=()=>reject(tx.error)});db.close();setJSONLocalSeguro("filaArquivosEnvio",await idsFilaArquivos())}
async function removerArquivoFila(id){try{const db=await abrirBancoUploads();await new Promise((resolve,reject)=>{const tx=db.transaction("arquivos","readwrite");tx.objectStore("arquivos").delete(id);tx.oncomplete=resolve;tx.onerror=()=>reject(tx.error)});db.close();setJSONLocalSeguro("filaArquivosEnvio",await idsFilaArquivos())}catch(e){}}
async function listarArquivosFila(){try{const db=await abrirBancoUploads();const itens=await new Promise((resolve,reject)=>{const tx=db.transaction("arquivos","readonly"),req=tx.objectStore("arquivos").getAll();req.onsuccess=()=>resolve(req.result||[]);req.onerror=()=>reject(req.error)});db.close();return itens.sort((a,b)=>(a.ordem||0)-(b.ordem||0)||String(a.id).localeCompare(String(b.id)))}catch(e){return[]}}
let filaProcessando=false;
async function processarFila(){if(filaProcessando||!navigator.onLine)return;const uploadUrl=(APP.UPLOAD_API_URL||"").trim();if(!uploadUrl)return;filaProcessando=true;try{const arquivos=await listarArquivosFila();for(const item of arquivos){try{const r=await postAPI(item,90000,uploadUrl);if(r?.timeout)throw new Error("Envio de arquivo ainda sem confirmação");await removerArquivoFila(item.id);if(item.tipo==="manifesto")setExecLocal(item.idExecucao,{...item,status:"finalizado",confirmacaoPendente:false})}catch(e){item.tentativas=(item.tentativas||0)+1;try{const db=await abrirBancoUploads();await new Promise((resolve,reject)=>{const tx=db.transaction("arquivos","readwrite");tx.objectStore("arquivos").put(item);tx.oncomplete=resolve;tx.onerror=()=>reject(tx.error)});db.close()}catch(_){}}}}finally{filaProcessando=false}}
window.addEventListener("online",processarFila);setInterval(processarFila,60000);
