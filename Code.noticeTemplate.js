function api_getNoticeTemplate(){
  try{
    const props = PropertiesService.getScriptProperties();
    const tpl = props.getProperty('NOTICE_TEMPLATE') || '';
    return { ok:true, template: tpl };
  }catch(e){
    return { ok:false, message: String(e) };
  }
}

function api_setNoticeTemplate(payload){
  try{
    const tpl = String(payload?.template || '').trim();
    if(!tpl) return { ok:false, message:'template is empty' };

    const props = PropertiesService.getScriptProperties();
    props.setProperty('NOTICE_TEMPLATE', tpl);

    return { ok:true };
  }catch(e){
    return { ok:false, message: String(e) };
  }
}
