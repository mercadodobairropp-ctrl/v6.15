// Painel Operacional Confiança v6.15
const PASTA_DRIVE_ID = "13xr39zvMhe-YtqWBNuUG9BsCz6tD7W9g";
const PLANILHA_ID = "1OViTBjCmDPs56dp2_g6vbIszFau3WNI1LahAbR29X94";
const TELEGRAM_TOKEN = "8414044142:AAHoof4NoOkqiM1FfeY9EmMekfodnqh0LN8";
const TELEGRAM_CHAT_ID = "5426828201";
const ADMIN_MESTRE = "01";
const SETUP_VERSION = "v6.15";
const ABAS={usuarios:["login","senha","nome","tipo","turnosPermitidos","empresasPermitidas"],turnos:["id","nome","empresaId","ativo"],empresas:["id","nome","turnos","ativo"],checklists:["id","nome","descricao","horario","horarioFim","turnos","dias","prioridade","responsaveisPermitidos","tarefas","ativo","empresaId"],execucoes:["idExecucao","data","idChecklist","nomeChecklist","horario","horarioFim","status","login","nomeUsuario","turno","urlPDF","urlFotos","criadoEm","nomeArquivo","reabertoPor","reabertoEm","novoHorarioFim"],logs:["dataHora","tipo","login","nomeUsuario","idChecklist","nomeChecklist","detalhe"]};
let SS_CACHE=null;
function doGet(e){try{e=e||{parameter:{}};setup_();const acao=e.parameter.acao||"status",cb=e.parameter.callback;if(acao==="getLogs"||acao==="log")manutencaoLogs_();let r={status:"ok",mensagem:"Servidor online v6.15"};if(acao==="getSyncInfo")r=syncInfo_();if(acao==="getBase")r=baseDados_(e.parameter.inicio,e.parameter.fim);if(acao==="getOperacional")r=operacionalDados_(e);if(acao==="getUsers")r={status:"ok",usuarios:listar_("usuarios")};if(acao==="getTurnos")r={status:"ok",turnos:listar_("turnos")};if(acao==="getEmpresas")r={status:"ok",empresas:listar_("empresas")};if(acao==="getChecklists")r={status:"ok",checklists:listar_("checklists")};if(acao==="getExecucoes")r={status:"ok",execucoes:listarExecucoes_(e.parameter.inicio,e.parameter.fim)};if(acao==="getLogs")r={status:"ok",logs:listarLogs_(e.parameter.inicio,e.parameter.fim,e.parameter.tipo)};if(acao==="log")r=logDados_(e);if(acao==="marcarEnvioPendente")r=marcarEnvioPendenteDados_(e);if(acao==="iniciarExecucao")r=iniciarExecucaoDados_(e);if(acao==="reabrirExecucao")r=reabrirExecucaoDados_(e);if(acao==="verificarVencidos")r=verificarChecklistsVencidos_(e.parameter.empresaId||"");const txt=JSON.stringify(r);if(cb)return ContentService.createTextOutput(cb+"("+txt+")").setMimeType(ContentService.MimeType.JAVASCRIPT);return ContentService.createTextOutput(txt).setMimeType(ContentService.MimeType.JSON)}catch(err){return ContentService.createTextOutput(JSON.stringify({status:"erro",mensagem:err.toString()})).setMimeType(ContentService.MimeType.JSON)}}
function doPost(e){try{e=e||{parameter:{}};setup_();const a=e.parameter.acao||"";if(a==="telegram")return enviarTelegram_(e.parameter.mensagem||"");if(a==="saveUser")return salvarUsuario_(e);if(a==="saveChecklist")return salvarChecklist_(e);if(a==="saveTurno")return salvarTurno_(e);if(a==="saveEmpresa")return salvarEmpresa_(e);if(a==="deleteEmpresa")return deleteEmpresa_(e);if(a==="deleteTurno")return deleteTurno_(e);if(a==="toggleChecklist")return toggleChecklist_(e);if(a==="deleteChecklist")return deleteChecklist_(e);if(a==="iniciarExecucao")return iniciarExecucao_(e);if(a==="cancelarExecucao")return cancelarExecucao_(e);if(a==="marcarEnvioPendente")return marcarEnvioPendente_(e);if(a==="confirmarUploadArquivos")return confirmarUploadArquivos_(e);if(a==="reabrirExecucao")return reabrirExecucao_(e);if(a==="log")return log_(e.parameter.tipo,e.parameter.login,e.parameter.nomeUsuario,e.parameter.idChecklist,e.parameter.nomeChecklist,e.parameter.detalhe);return ContentService.createTextOutput("Ação desconhecida: "+a)}catch(err){return ContentService.createTextOutput("ERRO: "+err.toString())}}
function testarInstalacaoPrincipal(){
  setup_();
  const info=syncInfo_();
  Logger.log(JSON.stringify(info));
  return info;
}
function ss_(){if(!SS_CACHE)SS_CACHE=SpreadsheetApp.openById(PLANILHA_ID);return SS_CACHE}
function sheet_(nome){return ss_().getSheetByName(nome)||aba_(nome)}
function aba_(nome){const ss=ss_();let sh=ss.getSheetByName(nome);const headers=ABAS[nome];if(!sh){sh=ss.insertSheet(nome);sh.getRange(1,1,1,headers.length).setValues([headers])}const atual=sh.getRange(1,1,1,Math.max(sh.getLastColumn(),headers.length)).getValues()[0].map(String);headers.forEach(h=>{if(atual.indexOf(h)===-1)sh.getRange(1,sh.getLastColumn()+1).setValue(h)});return sh}

function normalizarLogin_(login){
  login = String(login || "").trim();
  if (/^\d+$/.test(login)) return login.padStart(2, "0");
  return login;
}

function normalizarAbaUsuarios_(){
  const sh = aba_("usuarios");
  const values = sh.getDataRange().getValues();
  if (!values.length) return;

  const headers = values[0].map(String);
  const canonical = ABAS.usuarios;

  const hasDuplicateNome = headers.filter(h => h === "nome").length > 1;
  const headerErrado =
    headers.length !== canonical.length ||
    canonical.some((h, i) => headers[i] !== h) ||
    hasDuplicateNome;

  if (!headerErrado) return;

  const idx = {};
  headers.forEach((h, i) => {
    if (!idx[h]) idx[h] = [];
    idx[h].push(i);
  });

  const novos = [canonical];

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const loginRaw = row[idx.login ? idx.login[0] : 0];
    if (String(loginRaw || "").trim() === "") continue;

    const login = normalizarLogin_(loginRaw);
    const senha = String(row[idx.senha ? idx.senha[0] : 1] || "").trim();

    let nome = "";
    if (idx.nome) {
      for (const i of idx.nome) {
        const v = String(row[i] || "").trim();
        if (v) { nome = v; break; }
      }
    }
    if (!nome && login === ADMIN_MESTRE) nome = "Andre";

    const tipo = String(row[idx.tipo ? idx.tipo[0] : 3] || "operador").trim().toLowerCase();

    let turnosPermitidos = "";
    if (idx.turnosPermitidos) {
      turnosPermitidos = String(row[idx.turnosPermitidos[0]] || "").trim();
    }
    let empresasPermitidas = "";
    if (idx.empresasPermitidas) {
      empresasPermitidas = String(row[idx.empresasPermitidas[0]] || "").trim();
    }

    const existente = novos.findIndex(x => x[0] === login);
    const linha = [login, senha, nome, tipo, turnosPermitidos, empresasPermitidas];

    // login é a referência: se existir repetido, mantém a última linha preenchida
    if (existente >= 1) novos[existente] = linha;
    else novos.push(linha);
  }

  sh.clearContents();
  sh.getRange(1, 1, novos.length, canonical.length).setValues(novos);
}

function setup_(){
  const props=PropertiesService.getScriptProperties();
  if(props.getProperty("setup_version")===SETUP_VERSION)return;
  Object.keys(ABAS).forEach(aba_);
  seed_();
  normalizarAbaUsuarios_();
  props.setProperty("setup_version",SETUP_VERSION);
}
function seed_(){const u=aba_("usuarios");if(u.getLastRow()===1){u.appendRow(["01","7421","Andre","admin","todos","todos"]);u.appendRow(["02","7421","Katia","operador","",""]);u.appendRow(["04","1111","Maria","operador","",""]);u.appendRow(["08","0000","Tauna","operador","",""])}const t=aba_("turnos");if(t.getLastRow()===1){[["gerencial","Gerencial","","sim"]].forEach(r=>t.appendRow(r))}}
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
      if(!o.nome && o.login===ADMIN_MESTRE) o.nome="Andre";
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
function empresaPorTurno_(empresas){
  const mapa={};
  empresas.forEach(function(e){
    parseLista_(e.turnos).forEach(function(id){if(id&&!mapa[id])mapa[id]=String(e.id||"").trim()});
  });
  return mapa;
}
function completarVinculosEmpresa_(dados){
  const empresas=(dados.empresas||[]).filter(function(e){return String(e.ativo||"sim").toLowerCase()!=="nao"});
  const mapa=empresaPorTurno_(empresas);
  const turnos=(dados.turnos||[]).map(function(t){
    const o=Object.assign({},t);
    if(!String(o.empresaId||"").trim()&&mapa[o.id])o.empresaId=mapa[o.id];
    return o;
  });
  const checklists=(dados.checklists||[]).map(function(c){
    const o=Object.assign({},c);
    if(!String(o.empresaId||"").trim()){
      const turnos=parseLista_(o.turnos);
      for(var i=0;i<turnos.length;i++){
        if(mapa[turnos[i]]){o.empresaId=mapa[turnos[i]];break}
      }
    }
    if(!checklistCompleto_(o))o.ativo="nao";
    return o;
  });
  return {usuarios:dados.usuarios||[],turnos:turnos,empresas:empresas,checklists:checklists,execucoes:dados.execucoes||[]};
}
function baseDados_(inicio,fim){
  const dados=completarVinculosEmpresa_({
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
