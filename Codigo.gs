// Painel Operacional Confiança v6.15
const PLANILHA_ID = "1x21hmdwOodNuHTP1kFzW_UB7Li61VXWwiMBu1OCTpuI";
const ADMIN_MESTRE = "1";
const SETUP_VERSION = "v6.15-base-nova-admin1-20260623";
const ABAS={usuarios:["login","senha","nome","tipo","turnosPermitidos","empresasPermitidas"],turnos:["id","nome","empresaId","ativo"],empresas:["id","nome","ativo"],checklists:["id","nome","descricao","horario","horarioFim","turnos","dias","prioridade","tarefas","ativo","empresaId"],execucoes:["idExecucao","data","idChecklist","nomeChecklist","horario","horarioFim","status","login","nomeUsuario","turno","urlPDF","urlFotos","criadoEm","nomeArquivo","reabertoPor","reabertoEm","novoHorarioFim"],logs:["dataHora","tipo","login","nomeUsuario","idChecklist","nomeChecklist","detalhe"]};
let SS_CACHE=null;
function doGet(e){try{e=e||{parameter:{}};setup_();const acao=e.parameter.acao||"status",cb=e.parameter.callback;if(acao==="getLogs"||acao==="log")manutencaoLogs_();let r={status:"ok",mensagem:"Servidor online v6.15"};if(acao==="getSyncInfo")r=syncInfo_();if(acao==="getBase")r=baseDados_(e.parameter.inicio,e.parameter.fim);if(acao==="getOperacional")r=operacionalDados_(e);if(acao==="getUsers")r={status:"ok",usuarios:listar_("usuarios")};if(acao==="getTurnos")r={status:"ok",turnos:listar_("turnos")};if(acao==="getEmpresas")r={status:"ok",empresas:listar_("empresas")};if(acao==="getChecklists")r={status:"ok",checklists:listar_("checklists")};if(acao==="getExecucoes")r={status:"ok",execucoes:listarExecucoes_(e.parameter.inicio,e.parameter.fim)};if(acao==="getLogs")r={status:"ok",logs:listarLogs_(e.parameter.inicio,e.parameter.fim,e.parameter.tipo)};if(acao==="log")r=logDados_(e);if(acao==="marcarEnvioPendente")r=marcarEnvioPendenteDados_(e);if(acao==="iniciarExecucao")r=iniciarExecucaoDados_(e);if(acao==="reabrirExecucao")r=reabrirExecucaoDados_(e);if(acao==="verificarVencidos")r=verificarChecklistsVencidos_(e.parameter.empresaId||"");const txt=JSON.stringify(r);if(cb)return ContentService.createTextOutput(cb+"("+txt+")").setMimeType(ContentService.MimeType.JAVASCRIPT);return ContentService.createTextOutput(txt).setMimeType(ContentService.MimeType.JSON)}catch(err){return ContentService.createTextOutput(JSON.stringify({status:"erro",mensagem:err.toString()})).setMimeType(ContentService.MimeType.JSON)}}
function doPost(e){try{e=e||{parameter:{}};setup_();const a=e.parameter.acao||"";if(a==="telegram")return enviarTelegram_(e.parameter.mensagem||"");if(a==="saveUser")return salvarUsuario_(e);if(a==="saveChecklist")return salvarChecklist_(e);if(a==="saveTurno")return salvarTurno_(e);if(a==="saveEmpresa")return salvarEmpresa_(e);if(a==="deleteEmpresa")return deleteEmpresa_(e);if(a==="deleteTurno")return deleteTurno_(e);if(a==="toggleChecklist")return toggleChecklist_(e);if(a==="deleteChecklist")return deleteChecklist_(e);if(a==="iniciarExecucao")return iniciarExecucao_(e);if(a==="cancelarExecucao")return cancelarExecucao_(e);if(a==="marcarEnvioPendente")return marcarEnvioPendente_(e);if(a==="confirmarUploadArquivos")return confirmarUploadArquivos_(e);if(a==="reabrirExecucao")return reabrirExecucao_(e);if(a==="log")return log_(e.parameter.tipo,e.parameter.login,e.parameter.nomeUsuario,e.parameter.idChecklist,e.parameter.nomeChecklist,e.parameter.detalhe);return ContentService.createTextOutput("Ação desconhecida: "+a)}catch(err){return ContentService.createTextOutput("ERRO: "+err.toString())}}
function testarInstalacaoPrincipal(){
  setup_();
  const ss=ss_();
  const info=Object.assign(syncInfo_(),{
    planilhaId:ss.getId(),
    planilhaNome:ss.getName(),
    abas:ss.getSheets().map(function(sh){return sh.getName()})
  });
  Logger.log(JSON.stringify(info));
  return info;
}
function ss_(){if(!SS_CACHE)SS_CACHE=SpreadsheetApp.openById(PLANILHA_ID);return SS_CACHE}
function sheet_(nome){return ss_().getSheetByName(nome)||aba_(nome)}
function aba_(nome){const ss=ss_();let sh=ss.getSheetByName(nome);const headers=ABAS[nome];if(!sh){sh=ss.insertSheet(nome);sh.getRange(1,1,1,headers.length).setValues([headers])}const atual=sh.getRange(1,1,1,Math.max(sh.getLastColumn(),headers.length)).getValues()[0].map(String);headers.forEach(h=>{if(atual.indexOf(h)===-1)sh.getRange(1,sh.getLastColumn()+1).setValue(h)});return sh}

function normalizarLogin_(login){
  return String(login || "").trim();
}

function setup_(){
  const props=PropertiesService.getScriptProperties();
  if(props.getProperty("setup_version")===SETUP_VERSION)return;
  Object.keys(ABAS).forEach(aba_);
  criarAdminInicial_();
  removerAbaInicialVazia_();
  props.setProperty("setup_version",SETUP_VERSION);
}
function criarAdminInicial_(){
  const u=aba_("usuarios");
  if(u.getLastRow()>1)u.getRange(2,1,u.getLastRow()-1,u.getLastColumn()).clearContent();
  u.appendRow(["1","1111","André","admin","",""]);
}
function removerAbaInicialVazia_(){
  const ss=ss_();
  ss.getSheets().forEach(function(sh){
    if(ABAS[sh.getName()]||ss.getSheets().length<=1)return;
    if(sh.getLastRow()<=1&&sh.getLastColumn()<=1&&String(sh.getRange(1,1).getValue()||"").trim()==="")ss.deleteSheet(sh);
  });
}
function formatCell_(v,campo){
  if(Object.prototype.toString.call(v)==="[object Date]"){
    if(campo==="data") return Utilities.formatDate(v,Session.getScriptTimeZone(),"yyyy-MM-dd");
    if(campo==="criadoEm"||campo==="reabertoEm"||campo==="dataHora") return Utilities.formatDate(v,Session.getScriptTimeZone(),"yyyy-MM-dd'T'HH:mm:ss");
    return Utilities.formatDate(v,Session.getScriptTimeZone(),"HH:mm");
  }
  return String(v==null?"":v);
}
function listar_(nome){
  const sh=sheet_(nome),vals=sh.getDataRange().getValues(),h=vals[0].map(String),out=[];
  const mapaUsuarios = {};

  for(let i=1;i<vals.length;i++){
    if(vals[i].join("")==="")continue;
    let o={};
    h.forEach((k,j)=>{
      let valor=formatCell_(vals[i][j],k);
      if(nome==="usuarios" && k==="login") valor=normalizarLogin_(valor);
      o[k]=valor;
    });

    if(nome==="usuarios"){
      if(!o.nome && o.login===ADMIN_MESTRE) o.nome="André";
      o.tipo=String(o.tipo||"operador").toLowerCase();
      mapaUsuarios[o.login]=o;
    }else{
      out.push(o);
    }
  }

  if(nome==="usuarios") return Object.values(mapaUsuarios);
  return out;
}
function listarExecucoes_(inicio,fim){return listar_("execucoes").filter(r=>(!inicio||r.data>=inicio)&&(!fim||r.data<=fim))}
function listarLogs_(inicio,fim,tipo){
  return listar_("logs").filter(function(r){
    const dia=String(r.dataHora||"").slice(0,10);
    const tp=String(r.tipo||"");
    return (!inicio||dia>=inicio)&&(!fim||dia<=fim)&&(!tipo||tp.indexOf(tipo)===0);
  });
}
function revisionDados_(partes){
  const texto=JSON.stringify(partes);
  return Utilities.computeDigest(Utilities.DigestAlgorithm.MD5,texto).map(b=>("0"+((b+256)%256).toString(16)).slice(-2)).join("").slice(0,10);
}
function validarBase_(dados){
  const empresas=(dados.empresas||[]).filter(function(e){return String(e.ativo||"sim").toLowerCase()!=="nao"});
  const turnos=(dados.turnos||[]).filter(function(t){return String(t.empresaId||"").trim()&&String(t.ativo||"sim").toLowerCase()!=="nao"});
  const checklists=(dados.checklists||[]).map(function(c){
    const o=Object.assign({},c);
    if(!checklistCompleto_(o))o.ativo="nao";
    return o;
  });
  return {usuarios:dados.usuarios||[],turnos:turnos,empresas:empresas,checklists:checklists,execucoes:dados.execucoes||[]};
}
function baseDados_(inicio,fim){
  const dados=validarBase_({
    usuarios:listar_("usuarios"),
    turnos:listar_("turnos"),
    empresas:listar_("empresas"),
    checklists:listar_("checklists"),
    execucoes:listarExecucoes_(inicio,fim)
  });
  return {
    status:"ok",
    versao:"v6.15",
    revision:revisionDados_(dados),
    servidorEm:Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"yyyy-MM-dd'T'HH:mm:ss"),
    usuarios:dados.usuarios,
    turnos:dados.turnos,
    empresas:dados.empresas,
    checklists:dados.checklists,
    execucoes:dados.execucoes
  };
}
function diaSemanaData_(data){
  const s=String(data||dataHoje_()).slice(0,10);
  const d=new Date(s+"T12:00:00");
  return ["dom","seg","ter","qua","qui","sex","sab"][d.getDay()];
}
function usuarioPodeEmpresa_(u,empresaId){
  if(!u)return false;
  if(String(u.tipo||"").toLowerCase()==="admin")return true;
  const empresas=parseLista_(u.empresasPermitidas);
  return empresas.indexOf("todos")>=0||empresas.indexOf(String(empresaId||"").trim())>=0;
}
function usuarioPodeTurno_(u,turnoId){
  if(!u)return false;
  if(String(u.tipo||"").toLowerCase()==="admin")return true;
  const turnos=parseLista_(u.turnosPermitidos);
  return turnos.indexOf("todos")>=0||turnos.indexOf(String(turnoId||"").trim())>=0;
}
function operacionalDados_(e){
  const login=normalizarLogin_(e.parameter.login);
  const empresaId=String(e.parameter.empresaId||"").trim();
  const data=String(e.parameter.data||dataHoje_()).slice(0,10);
  const dia=diaSemanaData_(data);
  if(!login)return {status:"erro",motivo:"login_obrigatorio",mensagem:"Login não informado."};
  if(!empresaId)return {status:"erro",motivo:"empresa_dispositivo_nao_definida",mensagem:"Empresa do dispositivo não informada."};

  const dados=validarBase_({
    usuarios:listar_("usuarios"),
    turnos:listar_("turnos"),
    empresas:listar_("empresas"),
    checklists:listar_("checklists"),
    execucoes:listarExecucoes_(data,data)
  });
  const usuario=(dados.usuarios||[]).find(function(u){return normalizarLogin_(u.login)===login});
  const empresa=(dados.empresas||[]).find(function(emp){return String(emp.id||"").trim()===empresaId});
  if(!usuario)return {status:"erro",motivo:"usuario_nao_encontrado",mensagem:"Usuário não encontrado.",empresaId:empresaId,data:data};
  if(!empresa)return {status:"erro",motivo:"empresa_nao_encontrada",mensagem:"Empresa do dispositivo não encontrada.",usuario:usuario,empresaId:empresaId,data:data};
  if(!usuarioPodeEmpresa_(usuario,empresaId))return {status:"bloqueado",motivo:"empresa_nao_liberada",mensagem:"Empresa não liberada para usuário.",usuario:usuario,empresa:empresa,data:data};

  const turnosUsuario=(dados.turnos||[]).filter(function(t){
    return String(t.ativo||"sim").toLowerCase()!=="nao"&&String(t.empresaId||"").trim()===empresaId&&usuarioPodeTurno_(usuario,t.id);
  });
  const idsTurnos={};
  turnosUsuario.forEach(function(t){idsTurnos[String(t.id||"").trim()]=true});
  const checklists=(dados.checklists||[]).filter(function(c){
    if(String(c.ativo||"sim").toLowerCase()==="nao")return false;
    if(String(c.empresaId||"").trim()!==empresaId)return false;
    const dias=parseLista_(String(c.dias||"").toLowerCase());
    if(dias.length&&dias.indexOf(dia)<0)return false;
    return parseLista_(c.turnos).some(function(t){return idsTurnos[t]});
  });
  const idsChecks={};
  checklists.forEach(function(c){idsChecks[String(c.id||"").trim()]=true});
  const execucoes=(dados.execucoes||[]).filter(function(ex){return idsChecks[String(ex.idChecklist||"").trim()]});
  return {
    status:"ok",
    versao:"v6.15",
    servidorEm:Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"yyyy-MM-dd'T'HH:mm:ss"),
    data:data,
    dia:dia,
    usuario:usuario,
    empresa:empresa,
    turnos:turnosUsuario,
    checklists:checklists,
    execucoes:execucoes
  };
}
function syncInfo_(){
  const nomes=["usuarios","turnos","empresas","checklists","execucoes"];
  const partes=nomes.map(nome=>{
    const sh=aba_(nome);
    const lr=sh.getLastRow(),lc=sh.getLastColumn();
    const ultima=lr>1?sh.getRange(lr,1,1,lc).getDisplayValues()[0]:[];
    return nome+":"+lr+":"+lc+":"+JSON.stringify(ultima);
  }).join("|");
  const digest=Utilities.computeDigest(Utilities.DigestAlgorithm.MD5,partes).map(b=>("0"+((b+256)%256).toString(16)).slice(-2)).join("").slice(0,10);
  return {status:"ok",versao:"v6.15",revision:digest,servidorEm:Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"yyyy-MM-dd'T'HH:mm:ss")};
}
function salvarUsuario_(e){
  const login=normalizarLogin_(e.parameter.login);
  const ator=normalizarLogin_(e.parameter.atorLogin);
  if(login===ADMIN_MESTRE && ator!==ADMIN_MESTRE) return ContentService.createTextOutput("Admin mestre protegido");

  upsertObj_("usuarios","login",login,{
    login,
    senha:e.parameter.senha||"",
    nome:e.parameter.nome||"",
    tipo:e.parameter.tipo||"operador",
    turnosPermitidos:e.parameter.turnosPermitidos||"",
    empresasPermitidas:e.parameter.empresasPermitidas||""
  });

  log_("salvou_usuario",ator,"",login,e.parameter.nome||"","Usuário salvo/atualizado");
  return ContentService.createTextOutput("ok");
}

function normalizarHora_(valor){
  valor = String(valor || "").trim();
  if (/^\d{1,2}$/.test(valor)) {
    let h = Math.min(23, Math.max(0, parseInt(valor, 10)));
    return String(h).padStart(2, "0") + ":00";
  }
  let m = valor.match(/(\d{1,2}):?(\d{2})?/);
  if (m) {
    let h = Math.min(23, Math.max(0, parseInt(m[1] || "0", 10)));
    let min = Math.min(59, Math.max(0, parseInt(m[2] || "0", 10)));
    return String(h).padStart(2, "0") + ":" + String(min).padStart(2, "0");
  }
  return "00:00";
}

function parseLista_(v){
  return String(v||"").split(",").map(function(x){return x.trim()}).filter(Boolean);
}

function parseTarefas_(v){
  return String(v||"").replace(/\\n/g,"\n").split(/\n|;/).map(function(x){return x.trim()}).filter(Boolean);
}

function checklistCompleto_(c){
  return !!(
    String(c.id||"").trim() &&
    String(c.nome||"").trim() &&
    String(c.empresaId||"").trim() &&
    parseLista_(c.turnos).length &&
    parseTarefas_(c.tarefas).length &&
    normalizarHora_(c.horario)!=="00:00" &&
    normalizarHora_(c.horarioFim)!=="00:00"
  );
}


function salvarTurno_(e){
  const id=String(e.parameter.id||"").trim();
  const nome=String(e.parameter.nome||"").trim();
  const empresaId=String(e.parameter.empresaId||"").trim();
  const ativo=String(e.parameter.ativo||"sim").trim()||"sim";
  if(!id||!nome)return ContentService.createTextOutput("Dados incompletos");
  upsertObj_("turnos","id",id,{id,nome,empresaId,ativo});
  log_("salvou_turno",e.parameter.login,e.parameter.nomeUsuario,id,nome,"Turno salvo para empresa "+empresaId);
  return ContentService.createTextOutput("ok");
}
function deleteTurno_(e){
  const id=String(e.parameter.id||"").trim();
  deleteRow_("turnos","id",id);
  log_("excluiu_turno",e.parameter.login,e.parameter.nomeUsuario,id,"","Turno excluído");
  return ContentService.createTextOutput("ok");
}

function salvarEmpresa_(e){
  const id=String(e.parameter.id||"").trim();
  const nome=String(e.parameter.nome||"").trim();
  const ativo=String(e.parameter.ativo||"sim").trim()||"sim";
  if(!id||!nome)return ContentService.createTextOutput("Dados incompletos");
  upsertObj_("empresas","id",id,{id,nome,ativo});
  log_("salvou_empresa",e.parameter.login,e.parameter.nomeUsuario,id,nome,"Empresa salva");
  return ContentService.createTextOutput("ok");
}
function deleteEmpresa_(e){
  const id=String(e.parameter.id||"").trim();
  deleteRow_("empresas","id",id);
  log_("excluiu_empresa",e.parameter.login,e.parameter.nomeUsuario,id,"","Empresa excluída");
  return ContentService.createTextOutput("ok");
}

function salvarChecklist_(e){
  const obj={id:e.parameter.id,nome:e.parameter.nome,descricao:e.parameter.descricao,horario:normalizarHora_(e.parameter.horario),horarioFim:normalizarHora_(e.parameter.horarioFim),turnos:e.parameter.turnos,dias:e.parameter.dias,prioridade:e.parameter.prioridade,tarefas:e.parameter.tarefas,ativo:e.parameter.ativo,empresaId:e.parameter.empresaId||""};
  if(!checklistCompleto_(obj))obj.ativo="nao";
  upsertObj_("checklists","id",obj.id,obj);
  log_("salvou_checklist",e.parameter.login,e.parameter.nomeUsuario,obj.id,obj.nome,"Checklist salvo"+(obj.ativo==="nao"?" como inativo por cadastro incompleto":""));
  return ContentService.createTextOutput("ok");
}
function upsertObj_(abaNome,keyName,keyValue,obj){
  const sh=aba_(abaNome),vals=sh.getDataRange().getValues(),h=vals[0].map(String),keyCol=h.indexOf(keyName)+1,row=h.map(k=>obj[k]||"");
  const chaveNova = (abaNome==="usuarios" && keyName==="login") ? normalizarLogin_(keyValue) : String(keyValue).trim();

  for(let i=1;i<vals.length;i++){
    let chaveAtual = String(vals[i][keyCol-1]).trim();
    if(abaNome==="usuarios" && keyName==="login") chaveAtual = normalizarLogin_(chaveAtual);

    if(chaveAtual===chaveNova){
      sh.getRange(i+1,1,1,h.length).setValues([row]);
      return;
    }
  }
  sh.appendRow(row);
}
function toggleChecklist_(e){
  const c=checklistPorId_(e.parameter.id);
  let ativo=e.parameter.ativo;
  if(ativo==="sim"&&!checklistCompleto_(c))ativo="nao";
  updateField_("checklists","id",e.parameter.id,"ativo",ativo);
  log_("toggle_checklist",e.parameter.login,e.parameter.nomeUsuario,e.parameter.id,"","Ativo: "+ativo);
  return ContentService.createTextOutput("ok");
}
function deleteChecklist_(e){deleteRow_("checklists","id",e.parameter.id);log_("excluiu_checklist",e.parameter.login,e.parameter.nomeUsuario,e.parameter.id,"","Excluído");return ContentService.createTextOutput("ok")}
function updateField_(abaNome,keyName,keyValue,field,value){const sh=aba_(abaNome),vals=sh.getDataRange().getValues(),h=vals[0].map(String),keyCol=h.indexOf(keyName)+1,fCol=h.indexOf(field)+1;for(let i=1;i<vals.length;i++){if(String(vals[i][keyCol-1]).trim()===String(keyValue).trim()){sh.getRange(i+1,fCol).setValue(value);return}}}

function nomeUsuarioPorLogin_(login){
  login=String(login||"").trim();
  if(/^\d+$/.test(login)) login=login.padStart(2,"0");
  const usuarios=listar_("usuarios");
  const u=usuarios.find(x=>String(x.login)===login);
  if(u&&u.nome)return u.nome;
  if(login===ADMIN_MESTRE)return "André";
  return login?("Usuário "+login):"";
}

function deleteRow_(abaNome,keyName,keyValue){const sh=aba_(abaNome),vals=sh.getDataRange().getValues(),h=vals[0].map(String),keyCol=h.indexOf(keyName)+1;for(let i=1;i<vals.length;i++){if(String(vals[i][keyCol-1]).trim()===String(keyValue).trim()){sh.deleteRow(i+1);return}}}
function iniciarExecucaoDados_(e){
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(8000))return {status:"erro",mensagem:"Servidor ocupado. Tente novamente."};
  try{
    const p=e.parameter||{};
    const idExec=String(p.idExecucao||"").trim();
    const idChecklist=String(p.idChecklist||"").trim();
    const login=normalizarLogin_(p.login);
    if(!idExec||!idChecklist||!login)return {status:"erro",mensagem:"Dados incompletos para iniciar."};
    const existente=listar_("execucoes").find(function(r){return String(r.idExecucao||"").trim()===idExec})||{};
    const statusAtual=String(existente.status||"").toLowerCase();
    if(statusAtual==="executando"&&String(existente.login||"").trim()){
      return {status:"bloqueado",mensagem:"Checklist já está em execução.",execucao:existente};
    }
    if(statusAtual==="finalizado"||statusAtual==="aguardando_envio"){
      return {status:"bloqueado",mensagem:"Checklist já foi finalizado.",execucao:existente};
    }
    const execucao={
      idExecucao:idExec,
      data:p.data||dataHoje_(),
      idChecklist:idChecklist,
      nomeChecklist:p.nomeChecklist||"",
      horario:normalizarHora_(p.horario),
      horarioFim:normalizarHora_(p.horarioFim),
      status:"executando",
      login:login,
      nomeUsuario:p.nomeUsuario||nomeUsuarioPorLogin_(login),
      turno:p.turno||"",
      urlPDF:"",
      criadoEm:new Date().toISOString(),
      nomeArquivo:"",
      reabertoPor:existente.reabertoPor||"",
      reabertoEm:existente.reabertoEm||"",
      novoHorarioFim:existente.novoHorarioFim||""
    };
    upsertObj_("execucoes","idExecucao",idExec,execucao);
    log_("iniciou_execucao",login,execucao.nomeUsuario,idChecklist,p.nomeChecklist,"Iniciou "+idExec);
    return {status:"ok",execucao:execucao};
  }finally{
    lock.releaseLock();
  }
}
function iniciarExecucao_(e){
  const r=iniciarExecucaoDados_(e);
  return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
}
function cancelarExecucao_(e){
  const idExec=String(e.parameter.idExecucao||"").trim();
  if(!idExec)return ContentService.createTextOutput("Dados incompletos");
  const ex=listar_("execucoes").find(function(r){return String(r.idExecucao||"").trim()===idExec})||{};
  const st=String(ex.status||"").toLowerCase();
  if(st==="executando"||st==="reaberto")deleteRow_("execucoes","idExecucao",idExec);
  log_("cancelou_execucao",e.parameter.login,e.parameter.nomeUsuario,e.parameter.idChecklist||ex.idChecklist,e.parameter.nomeChecklist||ex.nomeChecklist,"Cancelou/liberou "+idExec);
  return ContentService.createTextOutput("ok");
}
function linhaExecucaoBase_(p,status,urlPDF,urlFotos){
  return {idExecucao:p.idExecucao,data:p.data||dataHoje_(),idChecklist:p.idChecklist,nomeChecklist:p.nomeChecklist,horario:normalizarHora_(p.horario),horarioFim:normalizarHora_(p.horarioFim),status:status,login:normalizarLogin_(p.login),nomeUsuario:p.nomeUsuario||nomeUsuarioPorLogin_(p.login),turno:p.turno||"",urlPDF:urlPDF||"",urlFotos:urlFotos||"",criadoEm:new Date().toISOString(),nomeArquivo:p.nomeArquivo||"",reabertoPor:"",reabertoEm:"",novoHorarioFim:""};
}
function marcarEnvioPendenteDados_(e){
  const p=e.parameter||{};
  if(!p.idExecucao||!p.idChecklist)return {status:"erro",mensagem:"Dados incompletos para marcar envio."};
  const linha=linhaExecucaoBase_(p,"aguardando_envio","","");
  upsertObj_("execucoes","idExecucao",p.idExecucao,linha);
  log_("aguardando_envio",p.login,p.nomeUsuario,p.idChecklist,p.nomeChecklist,"Arquivos serão enviados em segundo plano");
  return {status:"ok",execucao:linha};
}
function marcarEnvioPendente_(e){
  return ContentService.createTextOutput(JSON.stringify(marcarEnvioPendenteDados_(e))).setMimeType(ContentService.MimeType.JSON);
}
function confirmarUploadArquivos_(e){
  const p=e.parameter||{};
  if(!p.idExecucao)return ContentService.createTextOutput(JSON.stringify({status:"erro",mensagem:"Execução não informada"})).setMimeType(ContentService.MimeType.JSON);
  const linha=linhaExecucaoBase_(p,"finalizado",p.urlPDF,p.urlFotos);
  upsertObj_("execucoes","idExecucao",p.idExecucao,linha);
  log_("finalizou",p.login,p.nomeUsuario,p.idChecklist,p.nomeChecklist,"PDF: "+(p.urlPDF||"")+" | Fotos: "+(p.urlFotos||""));
  return ContentService.createTextOutput(JSON.stringify({status:"ok",execucao:linha})).setMimeType(ContentService.MimeType.JSON);
}
function checklistPorId_(id){
  const lista=listar_("checklists");
  return lista.find(c=>String(c.id).trim()===String(id).trim()) || {};
}
function dataHoje_(){
  return Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"yyyy-MM-dd");
}
function dataBR_(){
  return Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"dd/MM/yyyy");
}
function diaSemana_(){
  return ["dom","seg","ter","qua","qui","sex","sab"][new Date().getDay()];
}
function minutos_(h){
  const p=String(h||"00:00").slice(0,5).split(":");
  return (parseInt(p[0]||"0",10)*60)+(parseInt(p[1]||"0",10));
}
function agoraMinutos_(){
  const now=new Date();
  return now.getHours()*60+now.getMinutes();
}
function execId_(c,data){
  return data+"_"+c.id+"_"+String(c.horario).replace(":","-");
}
function nomeEmpresa_(id){
  const e=listar_("empresas").find(function(x){return String(x.id).trim()===String(id||"").trim()});
  return e&&e.nome?e.nome:(id||"Sem empresa");
}
function execFinalizada_(ex){
  if(!ex)return false;
  const st=String(ex.status||"").toLowerCase();
  return st==="finalizado"||st==="aguardando_envio";
}
function execEmAndamento_(ex){
  if(!ex)return false;
  return String(ex.status||"").toLowerCase()==="executando";
}
function execReabertaAindaValida_(ex){
  if(!ex||String(ex.status||"").toLowerCase()!=="reaberto")return false;
  return minutos_(ex.novoHorarioFim)>agoraMinutos_();
}
function telegramJaEnviado_(idExec){
  return PropertiesService.getScriptProperties().getProperty("telegram_vencido_"+idExec)==="1";
}
function marcarTelegramEnviado_(idExec){
  PropertiesService.getScriptProperties().setProperty("telegram_vencido_"+idExec,"1");
}
function mensagemChecklistVencido_(c){
  return [
    "Checklist não feito",
    "Empresa: "+nomeEmpresa_(c.empresaId),
    "Checklist: "+(c.nome||c.id),
    "Prazo: "+(c.horarioFim||"-"),
    "Data: "+dataBR_()
  ].join("\n");
}
function verificarChecklistsVencidos_(empresaId){
  const hoje=dataHoje_(),dia=diaSemana_(),agora=agoraMinutos_();
  const execs=listarExecucoes_(hoje,hoje);
  const porId={};
  execs.forEach(function(e){porId[String(e.idExecucao||"").trim()]=e});
  const enviados=[];

  listar_("checklists").forEach(function(c){
    if(String(c.ativo||"sim").toLowerCase()==="nao")return;
    if(!checklistCompleto_(c))return;
    if(empresaId&&String(c.empresaId||"").trim()!==String(empresaId).trim())return;
    const dias=parseLista_(String(c.dias||"").toLowerCase());
    if(dias.length&&dias.indexOf(dia)===-1)return;
    if(agora<=minutos_(c.horarioFim))return;

    const idExec=execId_(c,hoje);
    const ex=porId[idExec];
    if(execFinalizada_(ex)||execEmAndamento_(ex)||execReabertaAindaValida_(ex)||telegramJaEnviado_(idExec))return;

    const msg=mensagemChecklistVencido_(c);
    const envio=enviarTelegramMensagem_(msg);
    marcarTelegramEnviado_(idExec);
    enviados.push({idExecucao:idExec,idChecklist:c.id,nomeChecklist:c.nome,telegram:envio.codigo});
    log_("telegram_vencido","sistema","",c.id,c.nome,"Telegram vencido: "+envio.codigo+" "+envio.texto);
  });

  return {status:"ok",enviados:enviados.length,itens:enviados};
}
function verificarChecklistsVencidos(){
  return verificarChecklistsVencidos_("");
}
function instalarTriggerTelegramVencidos(){
  ScriptApp.getProjectTriggers().forEach(function(t){
    if(t.getHandlerFunction()==="verificarChecklistsVencidos")ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("verificarChecklistsVencidos").timeBased().everyMinutes(1).create();
}

function reabrirExecucaoDados_(e){
  const lock=LockService.getScriptLock();
  if(!lock.tryLock(8000))return {status:"erro",mensagem:"Servidor ocupado. Tente novamente."};
  try{
    const now=new Date();
    const fim=new Date(now.getTime()+60*60*1000);
    const novo=Utilities.formatDate(fim,Session.getScriptTimeZone(),"HH:mm");
    const reabertoEm=Utilities.formatDate(now,Session.getScriptTimeZone(),"yyyy-MM-dd'T'HH:mm:ss");
    const idExec=String(e.parameter.idExecucao||"").trim();
    const idChecklist=String(e.parameter.idChecklist||"").trim();
    if(!idExec||!idChecklist)return {status:"erro",mensagem:"Dados incompletos para reabrir."};

    const ex=listar_("execucoes").find(r=>String(r.idExecucao).trim()===idExec) || {};
    const c=checklistPorId_(idChecklist);

    const linha={
      idExecucao:idExec,
      data:ex.data||dataHoje_(),
      idChecklist:idChecklist,
      nomeChecklist:ex.nomeChecklist||c.nome||"",
      horario:ex.horario||c.horario||"",
      horarioFim:ex.horarioFim||c.horarioFim||"",
      status:"reaberto",
      login:"",
      nomeUsuario:"",
      turno:ex.turno||"",
      urlPDF:"",
      criadoEm:ex.criadoEm||new Date().toISOString(),
      nomeArquivo:"",
      reabertoPor:e.parameter.login||"",
      reabertoEm:reabertoEm,
      novoHorarioFim:novo
    };

    upsertObj_("execucoes","idExecucao",idExec,linha);
    log_("reabriu",e.parameter.login,e.parameter.nomeUsuario,idChecklist,linha.nomeChecklist,"Reaberto até "+novo);
    return {status:"ok",execucao:linha};
  }finally{
    lock.releaseLock();
  }
}
function reabrirExecucao_(e){
  const r=reabrirExecucaoDados_(e);
  return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
}
function log_(tipo,login,nomeUsuario,idChecklist,nomeChecklist,detalhe){aba_("logs").appendRow([new Date().toISOString(),tipo||"",login||"",nomeUsuario||"",idChecklist||"",nomeChecklist||"",detalhe||""]);return ContentService.createTextOutput("ok")}
function logDados_(e){
  const p=e.parameter||{};
  log_(p.tipo,p.login,p.nomeUsuario,p.idChecklist,p.nomeChecklist,p.detalhe);
  return {status:"ok"};
}
function manutencaoLogs_(){
  const props=PropertiesService.getScriptProperties();
  const hoje=dataHoje_();
  if(props.getProperty("logsManutencaoDia")===hoje)return;
  props.setProperty("logsManutencaoDia",hoje);
  const sh=aba_("logs"),lr=sh.getLastRow(),lc=sh.getLastColumn();
  if(lr<=1)return;
  const atual=hoje.slice(0,7),headers=ABAS.logs;
  const vals=sh.getRange(2,1,lr-1,lc).getValues();
  const grupos={},remover=[];
  vals.forEach(function(row,i){
    const raw=row[0],dt=raw instanceof Date?raw:new Date(raw);
    if(isNaN(dt.getTime()))return;
    const dia=Utilities.formatDate(dt,Session.getScriptTimeZone(),"yyyy-MM-dd");
    const mes=dia.slice(0,7);
    if(mes&&mes<atual){
      const abaNome="logs_"+mes.replace("-","_");
      if(!grupos[abaNome])grupos[abaNome]=[];
      grupos[abaNome].push(row);
      remover.push(i+2);
    }
  });
  Object.keys(grupos).forEach(function(nome){
    const ss=ss_();
    let arq=ss.getSheetByName(nome);
    if(!arq){arq=ss.insertSheet(nome);arq.appendRow(headers)}
    arq.getRange(arq.getLastRow()+1,1,grupos[nome].length,lc).setValues(grupos[nome]);
  });
  remover.reverse().forEach(function(row){sh.deleteRow(row)});
}
function enviarTelegramMensagem_(mensagem){
  const props=PropertiesService.getScriptProperties();
  const token=String(props.getProperty("TELEGRAM_TOKEN")||"").trim();
  const chatId=String(props.getProperty("TELEGRAM_CHAT_ID")||"").trim();
  if(!token||!chatId)throw new Error("Configure TELEGRAM_TOKEN e TELEGRAM_CHAT_ID nas Propriedades do Script.");
  const resp=UrlFetchApp.fetch("https://api.telegram.org/bot"+token+"/sendMessage",{
    method:"post",
    contentType:"application/json",
    payload:JSON.stringify({chat_id:chatId,text:mensagem}),
    muteHttpExceptions:true
  });
  return {codigo:resp.getResponseCode(),texto:resp.getContentText().slice(0,300)};
}
function enviarTelegram_(mensagem){
  const envio=enviarTelegramMensagem_(mensagem);
  log_("telegram_manual","sistema","","","",envio.codigo+" "+envio.texto);
  return ContentService.createTextOutput("Telegram: "+envio.codigo);
}
