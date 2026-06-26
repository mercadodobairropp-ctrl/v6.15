const PASTA_DRIVE_ID = "13xr39zvMhe-YtqWBNuUG9BsCz6tD7W9g";
const API_PRINCIPAL_URL = "https://script.google.com/macros/s/AKfycbw8efsu7_JEybMoxJbvt0gjn1yAHICY7HFSl0ZVl5OM3jl8iyWDk9Xz2_Zs7XgXD1Ir/exec";

function doGet(e){
  e=e||{parameter:{}};
  const cb=e.parameter.callback;
  const r={status:"ok",mensagem:"Upload online v6.15"};
  const txt=JSON.stringify(r);
  if(cb)return ContentService.createTextOutput(cb+"("+txt+")").setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(txt).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e){
  try{
    e=e||{parameter:{}};
    const a=e.parameter.acao||"";
    if(a==="uploadArquivoChecklist")return uploadArquivoChecklist_(e.parameter);
    return ContentService.createTextOutput(JSON.stringify({status:"erro",mensagem:"Ação desconhecida: "+a})).setMimeType(ContentService.MimeType.JSON);
  }catch(err){
    return ContentService.createTextOutput(JSON.stringify({status:"erro",mensagem:err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

function testarInstalacaoUpload(){
  const pasta=DriveApp.getFolderById(PASTA_DRIVE_ID);
  const r={status:"ok",mensagem:"Upload v6.15 com acesso ao Drive",pasta:pasta.getName(),url:pasta.getUrl()};
  Logger.log(JSON.stringify(r));
  return r;
}

function slugArquivo_(v){
  return String(v||"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/[\\/:*?"<>|]+/g," ").replace(/\s+/g," ").trim().slice(0,90)||"arquivo";
}

function subpasta_(pasta,nome){
  const it=pasta.getFoldersByName(nome);
  return it.hasNext()?it.next():pasta.createFolder(nome);
}

function pastaExecucao_(p){
  const raiz=DriveApp.getFolderById(PASTA_DRIVE_ID);
  const base=subpasta_(raiz,"Arquivos checklist");
  const data=subpasta_(base,String(p.data||"sem-data"));
  const checklist=subpasta_(data,slugArquivo_((p.idChecklist||"")+" - "+(p.nomeChecklist||"")));
  return subpasta_(checklist,slugArquivo_(p.idExecucao||new Date().toISOString()));
}

function salvarBlob_(pasta,nome,mime,base64){
  const blob=Utilities.newBlob(Utilities.base64Decode(base64),mime,nome);
  const arq=pasta.createFile(blob);
  return arq.getUrl();
}

function uploadArquivoChecklist_(p){
  const id=String(p.idExecucao||"").trim();
  const tipo=String(p.tipo||"").trim();
  if(!id)throw new Error("Execucao nao informada.");
  if(!["pdf","foto","manifesto"].includes(tipo))throw new Error("Tipo de arquivo invalido: "+tipo);
  if((tipo==="pdf"||tipo==="foto")&&!String(p.base64||"").trim())throw new Error("Arquivo vazio.");
  const pasta=pastaExecucao_(p);
  const props=PropertiesService.getScriptProperties();
  let url="";

  if(tipo==="pdf"){
    url=salvarBlob_(pasta,slugArquivo_(p.nomeArquivo||"checklist")+".pdf","application/pdf",p.base64||"");
    props.setProperty("pdf_"+id,url);
  }else if(tipo==="foto"){
    const fotos=subpasta_(pasta,"fotos");
    const idx=String(p.index||"0").padStart(2,"0");
    url=salvarBlob_(fotos,idx+" - "+slugArquivo_(p.tarefa||"foto")+".jpg","image/jpeg",p.base64||"");
  }else if(tipo==="manifesto"){
    const dados=JSON.parse(p.manifestoJson||"{}");
    const manifesto=Utilities.newBlob(JSON.stringify(dados,null,2),"application/json","manifesto.json");
    pasta.createFile(manifesto);
    const urlPDF=props.getProperty("pdf_"+id)||"";
    const urlFotos=pasta.getUrl();
    confirmarPrincipal_(p,urlPDF,urlFotos);
    props.deleteProperty("pdf_"+id);
    url=urlFotos;
  }

  return ContentService.createTextOutput(JSON.stringify({status:"ok",url:url,urlPasta:pasta.getUrl()})).setMimeType(ContentService.MimeType.JSON);
}

function confirmarPrincipal_(p,urlPDF,urlFotos){
  const resp=UrlFetchApp.fetch(API_PRINCIPAL_URL,{
    method:"post",
    muteHttpExceptions:true,
    payload:{
      acao:"confirmarUploadArquivos",
      idExecucao:p.idExecucao||"",
      data:p.data||"",
      idChecklist:p.idChecklist||"",
      nomeChecklist:p.nomeChecklist||"",
      horario:p.horario||"",
      horarioFim:p.horarioFim||"",
      login:p.login||"",
      nomeUsuario:p.nomeUsuario||"",
      turno:p.turno||"",
      nomeArquivo:p.nomeArquivo||"",
      urlPDF:urlPDF||"",
      urlFotos:urlFotos||""
    }
  });
  const codigo=resp.getResponseCode();
  const texto=resp.getContentText();
  let resposta=null;
  try{resposta=JSON.parse(texto)}catch(_){}
  if(codigo<200||codigo>=300||/^ERRO:/i.test(texto)||(resposta&&resposta.status!=="ok")){
    throw new Error("O script principal nao confirmou o upload. HTTP "+codigo+": "+texto.slice(0,200));
  }
}
