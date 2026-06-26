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

function dispositivoPorId_(deviceId){
  deviceId=String(deviceId||"").trim();
  if(!deviceId)return null;
  const lista=listar_("dispositivos");
  return lista.find(function(d){return String(d.deviceId||"").trim()===deviceId})||null;
}
function getDispositivoDados_(e){
  const deviceId=String(e.parameter.deviceId||"").trim();
  const d=dispositivoPorId_(deviceId);
  if(!deviceId)return {status:"erro",mensagem:"ID do dispositivo não informado."};
  if(!d||String(d.ativo||"sim").toLowerCase()==="nao")return {status:"ok",deviceId:deviceId,empresaId:"",dispositivo:null};
  return {status:"ok",deviceId:deviceId,empresaId:String(d.empresaId||"").trim(),dispositivo:d};
}
function salvarDispositivoDados_(e){
  const deviceId=String(e.parameter.deviceId||"").trim();
  const empresaId=String(e.parameter.empresaId||"").trim();
  const nome=String(e.parameter.nome||"").trim();
  if(!deviceId)return {status:"erro",mensagem:"ID do dispositivo não informado."};
  const obj={deviceId:deviceId,empresaId:empresaId,nome:nome,ativo:"sim",atualizadoEm:new Date().toISOString()};
  upsertObj_("dispositivos","deviceId",deviceId,obj);
  log_("salvou_dispositivo",e.parameter.login,e.parameter.nomeUsuario,deviceId,"","Empresa do dispositivo: "+empresaId);
  return {status:"ok",dispositivo:obj};
}
function salvarDispositivo_(e){
  return ContentService.createTextOutput(JSON.stringify(salvarDispositivoDados_(e))).setMimeType(ContentService.MimeType.JSON);
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
