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
