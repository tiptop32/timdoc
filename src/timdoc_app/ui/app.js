const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => [...document.querySelectorAll(selector)];
const state = { act: null, customer: null, settings: null, outputDirectory: "" };

const api = async (method, ...args) => {
  if (!window.pywebview?.api?.[method]) throw new Error("Откройте приложение Timdoc, а не HTML-файл в браузере");
  const response = await window.pywebview.api[method](...args);
  if (!response.ok) throw new Error(response.error || "Операция не выполнена");
  return response.data ?? response;
};
const busy = (visible) => $("#busy").classList.toggle("hidden", !visible);
const fail = (error) => { const box=$("#toast"); box.textContent=error.message || String(error); box.classList.remove("hidden"); setTimeout(()=>box.classList.add("hidden"),7000); };
const setStep = (number) => $$(".step").forEach((item,index)=>item.classList.toggle("active",index < number));
const customerFields = {
  customer_name:"name", customer_inn:"inn", customer_address:"address", customer_country:"country",
  customer_postal_code:"postal_code", customer_region:"region", customer_district:"district",
  customer_locality:"locality", customer_street:"street", customer_house:"house",
  representative_name:"representative_name", representative_position:"representative_position",
  customer_phone:"phone", customer_email:"email"
};
function writeCustomer(customer){ Object.entries(customerFields).forEach(([name,key])=>{ document.querySelector(`[name="${name}"]`).value=customer[key] || ""; }); }
function readCustomer(){ const result={}; Object.entries(customerFields).forEach(([name,key])=>result[key]=document.querySelector(`[name="${name}"]`).value.trim()); return result; }
function writeSettings(settings){ Object.keys(settings).forEach(key=>{ const field=document.querySelector(`#settings-form [name="${key}"]`); if(field) field.value=settings[key] || ""; }); }
function readSettings(){ return Object.fromEntries([...$("#settings-form").elements].filter(x=>x.name).map(x=>[x.name,x.value.trim()])); }
function equipmentRow(item={name:"",serial_number:"",manufacture_year:""}){
  const row=document.createElement("tr"); row.innerHTML=`<td><input aria-label="Модель" data-key="name"></td><td><input aria-label="Заводской номер" data-key="serial_number"></td><td><input aria-label="Год выпуска" data-key="manufacture_year"></td><td><button class="remove" type="button" aria-label="Удалить строку">×</button></td>`;
  row.querySelectorAll("input").forEach(input=>input.value=item[input.dataset.key] || ""); row.querySelector(".remove").onclick=()=>row.remove(); return row;
}
function writeEquipment(items){ const body=$("#equipment-body"); body.innerHTML=""; items.forEach(item=>body.append(equipmentRow(item))); }
function readEquipment(){ return $$("#equipment-body tr").map(row=>Object.fromEntries([...row.querySelectorAll("input")].map(input=>[input.dataset.key,input.value.trim()]))); }
async function initialize(){ try { const data=await api("bootstrap"); state.settings=data.settings; writeSettings(state.settings); } catch(error){ fail(error); } }
$("#choose-pdf").onclick=async()=>{ try { const selected=await api("choose_pdf"); if(!selected.path)return; $("#pdf-path").textContent=selected.path; busy(true); const data=await api("analyze_act",selected.path); state.act=data.act; state.customer=data.customer; writeCustomer(data.customer); writeEquipment(data.act.equipment); $("#review-summary").textContent=`${data.act.equipment.length} ед. техники, будет создано ${data.act.equipment.length*3} документа`; const warnings=$("#warnings"); warnings.textContent=data.act.warnings.join(" · "); warnings.classList.toggle("hidden",!data.act.warnings.length); $("#stage-upload").classList.add("hidden"); $("#stage-review").classList.remove("hidden"); setStep(2); } catch(error){fail(error)} finally{busy(false)} };
$("#add-equipment").onclick=()=>$("#equipment-body").append(equipmentRow());
$("#choose-output").onclick=async()=>{ try{const selected=await api("choose_folder"); if(selected.path){state.outputDirectory=selected.path;$("#output-path").value=selected.path}}catch(error){fail(error)} };
$("#settings-open").onclick=()=>$("#settings-dialog").showModal();
$("#choose-templates").onclick=async()=>{ try{const selected=await api("choose_folder"); if(!selected.path)return; busy(true); const data=await api("prepare_templates",selected.path); state.settings=data.settings; writeSettings(state.settings); }catch(error){fail(error)}finally{busy(false)} };
$("#save-settings").onclick=async()=>{ try{state.settings=await api("save_settings",readSettings()); $("#settings-dialog").close();}catch(error){fail(error)} };
$("#generate").onclick=async()=>{ try{ if(!state.outputDirectory)throw new Error("Выберите папку сохранения"); state.settings=readSettings(); busy(true); const result=await api("generate",{customer:readCustomer(),settings:state.settings,equipment:readEquipment(),document_date:state.act.document_date,output_directory:state.outputDirectory}); $("#stage-review").classList.add("hidden"); $("#stage-result").classList.remove("hidden"); $("#result-summary").textContent=`Создано ${result.files.length} DOCX. По три файла на каждую единицу техники.`; $("#archive-path").textContent=result.archive_path; setStep(3); }catch(error){fail(error)}finally{busy(false)} };
$("#new-act").onclick=()=>location.reload();
window.addEventListener("pywebviewready",initialize);
