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
