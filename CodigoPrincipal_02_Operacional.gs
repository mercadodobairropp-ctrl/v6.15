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
  const nomes=["usuarios","turnos","empresas","dispositivos","checklists","execucoes"];
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


