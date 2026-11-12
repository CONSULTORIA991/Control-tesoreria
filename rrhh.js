"use strict";
/* =====================================================================
   Módulo de Recursos Humanos
   Empleados · Novedades · Nómina quincenal · Seguridad social (Asopagos)
   · Contratos y documentos · Ajustes.
   Usa las utilidades de index.html (h, S, fmt, abrirModal, VIEWS…).
   Los datos viven en la tabla docs con colecciones "rh_…": solo las ven
   la administradora y el rol Recursos humanos de cada empresa (RLS).
   ===================================================================== */
(function(){
const XLSX_URL="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
const RH={on:false,emp:[],nov:[],nomina:{},planillas:[],docs:[],empresa:{},param:{},plantillas:{},ok:{emp:false,nov:false,cfg:false}};
window.RH=RH;

/* Valores legales por defecto (Colombia). Se cambian en Ajustes cada enero. */
const PARAM_BASE={smmlv:1750905,auxTransporte:249095,horasMes:210,recargoDominical:90};
const P=()=>{const p={...PARAM_BASE};for(const k of Object.keys(PARAM_BASE)){const v=Number(RH.param[k]);if(v>0)p[k]=v}return p};
const puedeRH=()=>["admin","rrhh"].includes(miRol());

function iniciar(){
  if(RH.on||!S.db||!puedeRH())return;RH.on=true;
  const sub=(col,fn)=>S.db.collection(col).onSnapshot(s=>{fn(s.docs.map(d=>({id:d.id,...d.data()})));refrescar()},e=>console.warn(col,e));
  sub("rh_empleados",l=>{RH.emp=l;RH.ok.emp=true});
  sub("rh_novedades",l=>{RH.nov=l;RH.ok.nov=true});
  sub("rh_nomina",l=>{RH.nomina=Object.fromEntries(l.map(x=>[x.id,x]))});
  sub("rh_planillas",l=>{RH.planillas=l});
  sub("rh_documentos",l=>{RH.docs=l});
  sub("rh_config",l=>{const m=Object.fromEntries(l.map(x=>[x.id,x]));RH.empresa=m.empresa||{};RH.param=m.parametros||{};RH.ok.cfg=true});
  sub("rh_plantillas",l=>{RH.plantillas=Object.fromEntries(l.map(x=>[x.id,x]))});
}
window.rhIniciar=iniciar;

document.head.append(h("style",null,`
.papel{background:#fff;color:#111;border:1px solid var(--line);border-radius:8px;padding:26px 30px;max-height:58vh;overflow:auto}
.doc-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(250px,1fr));gap:12px}
.doc-card{display:flex;flex-direction:column;gap:8px;align-items:flex-start}
.doc-card p{margin:0;flex:1}
.campos-chips{display:flex;flex-wrap:wrap;gap:6px}
.campos-chips button{font:500 .78rem var(--f-num);padding:3px 8px;border-radius:6px}
.sec-form{border-top:1px solid var(--line);padding-top:12px;margin-top:4px}
.tachado td{color:var(--muted);text-decoration:line-through}
textarea.plantilla{min-height:340px;font:.86rem/1.5 var(--f-num)}
`));

/* ---------- utilidades ---------- */
const TIPOS_DOC={CC:"cédula de ciudadanía",CE:"cédula de extranjería",TI:"tarjeta de identidad",PA:"pasaporte",PT:"permiso por protección temporal"};
const CONTRATO_LBL={indefinido:"Término indefinido",fijo:"Término fijo",obra_labor:"Obra o labor"};
const CONTRATO_TXT={indefinido:"a término indefinido",fijo:"a término fijo",obra_labor:"por duración de la obra o labor"};
const nombreEmp=e=>e?[e.primerNombre,e.segundoNombre,e.primerApellido,e.segundoApellido].filter(Boolean).join(" "):"";
const empPor=id=>RH.emp.find(e=>e.id===id);
const porNombre=(a,b)=>nombreEmp(a).localeCompare(nombreEmp(b),"es");
const soloNum=s=>String(s||"").replace(/\D/g,"");
const numF=new Intl.NumberFormat("es-CO",{maximumFractionDigits:0});
const MESES=["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"];
const fechaLarga=f=>f?`${+f.slice(8,10)} de ${MESES[+f.slice(5,7)-1]} de ${f.slice(0,4)}`:"";
const ultimoDia=m=>{const [y,mo]=m.split("-").map(Number);return new Date(Date.UTC(y,mo,0)).getUTCDate()};
/* día comercial (meses de 30 días): el último día del mes cuenta como 30 */
const cd=f=>{const d=+f.slice(8,10);return d>=ultimoDia(f.slice(0,7))?30:Math.min(d,30)};
function diasCom(desde,hasta,ini,fin){const a=desde>ini?desde:ini,b=hasta<fin?hasta:fin;if(!a||!b||a>b)return 0;return cd(b)-cd(a)+1}
function salarioEn(e,fecha){let s=Number(e.salario)||0,f="";for(const n of RH.nov)if(n.empleadoId===e.id&&!n.anulado&&n.tipo==="cambio_salario"&&n.desde<=fecha&&n.desde>=f&&Number(n.valor)>0){f=n.desde;s=Number(n.valor)}return s}
const activoEn=(e,desde,hasta)=>!!e.fechaIngreso&&e.fechaIngreso<=hasta&&(!e.fechaRetiro||e.fechaRetiro>=desde);
function estadoEmp(e){const hf=hoy();if(e.fechaRetiro&&e.fechaRetiro<hf)return {k:"retirado",chip:h("span",{class:"chip neu"},"Retirado")};
  if(e.fechaRetiro)return {k:"activo",chip:h("span",{class:"chip warn"},"Se retira el "+fmtF(e.fechaRetiro))};
  return {k:"activo",chip:h("span",{class:"chip ok"},"Activo")}}
function letras(n){
  n=Math.floor(Math.abs(n));if(!n)return "cero";
  const U=["","uno","dos","tres","cuatro","cinco","seis","siete","ocho","nueve","diez","once","doce","trece","catorce","quince","dieciséis","diecisiete","dieciocho","diecinueve","veinte","veintiuno","veintidós","veintitrés","veinticuatro","veinticinco","veintiséis","veintisiete","veintiocho","veintinueve"];
  const D=["","","","treinta","cuarenta","cincuenta","sesenta","setenta","ochenta","noventa"];
  const C=["","ciento","doscientos","trescientos","cuatrocientos","quinientos","seiscientos","setecientos","ochocientos","novecientos"];
  const cien=x=>{if(x===100)return "cien";const r=x%100;let s=C[Math.floor(x/100)];if(r)s+=(s?" ":"")+(r<30?U[r]:D[Math.floor(r/10)]+(r%10?" y "+U[r%10]:""));return s};
  const apoc=s=>s.replace(/veintiuno$/,"veintiún").replace(/uno$/,"un");
  const mil=x=>{const m=Math.floor(x/1000),r=x%1000;let s=m?(m===1?"mil":apoc(cien(m))+" mil"):"";if(r)s+=(s?" ":"")+cien(r);return s};
  const mm=Math.floor(n/1e6),r=n%1e6;let s=mm?(mm===1?"un millón":apoc(mil(mm))+" millones"):"";
  if(r)s+=(s?" ":"")+mil(r);else if(mm)s+=" de";return s;
}
const pesosLetras=v=>(letras(v)+" pesos m/cte").toUpperCase();
function duracion(a,b){if(!a||!b)return "";const d=diasEntre(a,b)+1;
  for(let m=1;m<=48;m++){const y=+a.slice(0,4),mo=+a.slice(5,7)-1+m,dd=+a.slice(8,10);const t=new Date(Date.UTC(y,mo,dd)-86400000).toISOString().slice(0,10);if(t===b)return `${letras(m)} (${m}) ${m===1?"mes":"meses"}`}
  return `${letras(d)} (${d}) días`}
async function guardarArchivo(nombre,blob){if(!S.downloads){toast("La descarga no está disponible.");return}try{await S.downloads.save({filename:nombre,data:blob})}catch(e){toast("No se pudo descargar el archivo.")}}
async function excel(nombre,hojas){await cargarScript(XLSX_URL);const wb=XLSX.utils.book_new();
  for(const [titulo,filas] of hojas){const ws=XLSX.utils.aoa_to_sheet(filas);ws["!cols"]=filas[0].map((c,i)=>({wch:Math.min(40,Math.max(10,...filas.slice(0,60).map(r=>String(r[i]??"").length))+2)}));XLSX.utils.book_append_sheet(wb,ws,titulo)}
  await guardarArchivo(nombre,new Blob([XLSX.write(wb,{type:"array",bookType:"xlsx"})],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}))}
function selEmp(valor,soloActivos,onchange,id){
  const l=RH.emp.filter(e=>!soloActivos||estadoEmp(e).k==="activo"||e.id===valor).sort(porNombre);
  return h("select",{id:id||"rh-sel-emp",onchange:onchange?e=>onchange(e.target.value):null},h("option",{value:""},l.length?"Elige el empleado…":"Primero crea un empleado"),l.map(e=>h("option",{value:e.id,selected:e.id===valor?true:null},`${nombreEmp(e)} · ${e.documento}`)));
}
const campo=(l,el,req,ayuda)=>h("label",{class:"f"},h("span",null,l,req?h("span",{class:"req"}," *"):null),el,ayuda?h("span",{class:"small muted",style:"font-weight:400"},ayuda):null);
const kpi=(lab,val,sub,alerta)=>h("div",{class:"kpi"+(alerta?" alert":"")},h("span",{class:"lab"},lab),h("span",{class:"val"},val),sub?h("span",{class:"sub"},sub):null);
const cargandoRH=()=>!RH.ok.emp||!RH.ok.nov;

/* ---------- novedades: qué son y cómo afectan la nómina ---------- */
const NOV={
  vacaciones:{l:"Vacaciones",g:"Ausencias pagadas",dias:1,pago:1,pila:"VAC",ayuda:"Días de vacaciones disfrutadas. Se pagan con el salario completo."},
  licencia_remunerada:{l:"Licencia remunerada (luto, calamidad, permiso pagado)",g:"Ausencias pagadas",dias:1,pago:1,pila:"VAC",ayuda:"El empleado no trabaja pero recibe su salario completo."},
  licencia_maternidad:{l:"Licencia de maternidad o paternidad",g:"Ausencias pagadas",dias:1,pago:1,pila:"LMA",ayuda:"Se paga el salario completo; la EPS lo reconoce después a la empresa."},
  incapacidad_general:{l:"Incapacidad por enfermedad",g:"Incapacidades",dias:1,pago:"ige",pila:"IGE",ayuda:"Se paga al 66,67 % del salario, sin bajar del salario mínimo diario."},
  incapacidad_laboral:{l:"Incapacidad por accidente de trabajo",g:"Incapacidades",dias:1,pago:1,pila:"IRL",ayuda:"Se paga al 100 %; la ARL la reconoce a la empresa."},
  licencia_no_remunerada:{l:"Licencia no remunerada",g:"Ausencias no pagadas",dias:1,pago:0,pila:"SLN",ayuda:"Permiso sin pago: esos días no se pagan."},
  suspension:{l:"Suspensión del contrato",g:"Ausencias no pagadas",dias:1,pago:0,pila:"SLN",doc:"suspension",ayuda:"Esos días no se pagan. Al guardar puedes generar la carta de suspensión."},
  ausencia:{l:"Ausencia sin justificar",g:"Ausencias no pagadas",dias:1,pago:0,pila:"SLN",ayuda:"Días que el empleado faltó sin permiso: no se pagan."},
  horas_extra:{l:"Horas extra y recargos",g:"Pagos adicionales",horas:1,salarial:1,ayuda:"Escribe cuántas horas y de qué tipo. El valor se calcula solo con el salario del empleado."},
  comision:{l:"Comisión o bonificación salarial",g:"Pagos adicionales",valor:1,salarial:1,ayuda:"Suma al pago y también a la base de salud y pensión."},
  bonificacion:{l:"Bonificación o auxilio no salarial",g:"Pagos adicionales",valor:1,ayuda:"Suma al pago, pero no hace parte de la base de salud y pensión."},
  prestamo:{l:"Descuento por préstamo o anticipo",g:"Descuentos",valor:1,descuento:1,ayuda:"Se resta del pago de la quincena de esa fecha."},
  libranza:{l:"Descuento por libranza o embargo",g:"Descuentos",valor:1,descuento:1,ayuda:"Se resta del pago de la quincena de esa fecha."},
  retencion:{l:"Retención en la fuente",g:"Descuentos",valor:1,descuento:1,ayuda:"Valor de retención calculado por contabilidad."},
  otro_descuento:{l:"Otro descuento",g:"Descuentos",valor:1,descuento:1,ayuda:"Se resta del pago de la quincena de esa fecha."},
  cambio_salario:{l:"Cambio de salario",g:"Contrato",valor:1,pila:"VSP",ayuda:"Escribe el nuevo salario mensual y desde qué fecha rige."},
};
const HORAS={HED:["Hora extra diurna",()=>1.25],HEN:["Hora extra nocturna",()=>1.75],RN:["Recargo nocturno",()=>0.35],
  RDF:["Recargo dominical o festivo",p=>p.recargoDominical/100],HEDDF:["Hora extra diurna en domingo o festivo",p=>1.25+p.recargoDominical/100],
  HENDF:["Hora extra nocturna en domingo o festivo",p=>1.75+p.recargoDominical/100],RNDF:["Recargo nocturno en domingo o festivo",p=>0.35+p.recargoDominical/100]};
const valorHoras=(n,sal,p)=>Math.round((Number(n.horas)||0)*sal/p.horasMes*(HORAS[n.subtipo]?.[1](p)||0));
function novTexto(n){const t=NOV[n.tipo]||{l:n.tipo};
  if(t.dias)return {cuando:n.hasta&&n.hasta!==n.desde?`${fmtF(n.desde)} a ${fmtF(n.hasta)}`:fmtF(n.desde),cuanto:`${diasEntre(n.desde,n.hasta||n.desde)+1} días`};
  if(t.horas)return {cuando:fmtF(n.desde),cuanto:`${n.horas} h · ${HORAS[n.subtipo]?.[0]||""}${Number(n.valor)>0?" · "+fmt(n.valor):""}`};
  return {cuando:(n.tipo==="cambio_salario"?"Desde ":"")+fmtF(n.desde),cuanto:fmt(n.valor)}}

/* ---------- nómina quincenal ---------- */
function periodo(mes,q){return {mes,q,id:`${mes}-Q${q}`,desde:q===1?`${mes}-01`:`${mes}-16`,hasta:q===1?`${mes}-15`:`${mes}-${pad(ultimoDia(mes))}`,
  nombre:`${q===1?"Primera":"Segunda"} quincena de ${mesNombre(mes).toLowerCase()}`}}
function liquidar(e,per,p){
  const ini=per.desde,fin=per.hasta;if(!activoEn(e,ini,fin))return null;
  const sal=salarioEn(e,fin),dia=sal/30,dc=diasCom(e.fechaIngreso,e.fechaRetiro||"9999-12-31",ini,fin);
  let dPag=0,dIge=0,dNo=0,extras=0,comis=0,noSal=0,desc=0,prest=0,hExtra=0,hDom=0,vDom=0;const det=[];
  for(const n of RH.nov){
    if(n.empleadoId!==e.id||n.anulado)continue;const t=NOV[n.tipo];if(!t||n.tipo==="cambio_salario")continue;
    if(t.dias){const d=diasCom(n.desde,n.hasta||n.desde,ini,fin);if(!d)continue;
      if(t.pago==="ige")dIge+=d;else if(t.pago)dPag+=d;else dNo+=d;det.push(`${t.l}: ${d} ${d===1?"día":"días"}`)}
    else{if(!(n.desde>=ini&&n.desde<=fin))continue;const v=n.tipo==="horas_extra"&&!(Number(n.valor)>0)?valorHoras(n,sal,p):Number(n.valor)||0;
      if(t.descuento){desc+=v;if(n.tipo==="prestamo")prest+=v}
      else if(n.tipo==="horas_extra"){extras+=v;if(n.subtipo==="RDF"||n.subtipo==="RNDF"){hDom+=Number(n.horas)||0;vDom+=v}else hExtra+=Number(n.horas)||0}
      else if(t.salarial)comis+=v;else noSal+=v;det.push(`${t.l}: ${fmt(v)}`)}
  }
  const dTrab=Math.max(0,dc-dPag-dIge-dNo),r=x=>Math.round(Math.round(x*100)/100); /* redondeo al peso sin errores de decimales */
  const L={empleadoId:e.id,nombre:nombreEmp(e),documento:e.documento||"",cargo:e.cargo||"",salario:sal,dias:dc,diasTrab:dTrab,diasPagados:dPag,diasIncap:dIge,diasNoPagados:dNo,
    basico:r(dia*dTrab),ausPag:r(dia*dPag),incap:r(Math.min(dia,Math.max(dia*2/3,p.smmlv/30))*dIge),aux:sal<=2*p.smmlv?r(p.auxTransporte/30*dTrab):0,
    extras:r(extras),comis:r(comis),noSal:r(noSal),desc:r(desc),prest:r(prest),hExtra,hDom,vDom:r(vDom),detalle:det};
  /* Los totales se calculan con los valores exactos y se redondean una sola vez al final,
     para que el neto coincida con lo que se paga (cada renglón se muestra redondeado al peso). */
  const auxX=sal<=2*p.smmlv?p.auxTransporte/30*dTrab:0,ibc=dia*dTrab+dia*dPag+Math.min(dia,Math.max(dia*2/3,p.smmlv/30))*dIge+extras+comis,k=sal/p.smmlv;
  const fspX=ibc*(k<4?0:k<16?0.01:k<17?0.012:k<18?0.014:k<19?0.016:k<20?0.018:0.02);
  L.salud=r(ibc*0.04);L.pension=r(ibc*0.04);L.fsp=r(fspX);
  L.devengado=r(ibc+auxX+noSal);L.deducciones=r(ibc*0.08+fspX+desc);L.neto=L.devengado-L.deducciones;
  return L;
}
function nominaDe(per){const doc=RH.nomina[per.id];if(doc?.estado==="cerrada")return {cerrada:true,doc,filas:doc.filas||[]};
  const p=P();return {cerrada:false,doc,filas:RH.emp.slice().sort(porNombre).map(e=>liquidar(e,per,p)).filter(Boolean)}}
const suma=(l,k)=>l.reduce((s,x)=>s+(Number(x[k])||0),0);

/* ---------- seguridad social: datos del mes para la planilla (PILA) ---------- */
function pilaMes(mes){
  const ini=`${mes}-01`,fin=`${mes}-${pad(ultimoDia(mes))}`;
  return RH.emp.filter(e=>activoEn(e,ini,fin)).sort(porNombre).map(e=>{
    const dc=diasCom(e.fechaIngreso,e.fechaRetiro||"9999-12-31",ini,fin),sal=salarioEn(e,fin),nov={};
    const add=(cod,d,a,b)=>{const x=nov[cod]??={dias:0,ini:a,fin:b};x.dias+=d;if(a<x.ini)x.ini=a;if(b>x.fin)x.fin=b};
    for(const n of RH.nov){if(n.empleadoId!==e.id||n.anulado)continue;const t=NOV[n.tipo];if(!t?.pila)continue;
      if(t.dias){const b0=n.hasta||n.desde,a=n.desde>ini?n.desde:ini,b=b0<fin?b0:fin;if(a>b)continue;add(t.pila,diasCom(n.desde,b0,ini,fin),a,b)}
      else if(n.desde>=ini&&n.desde<=fin)add(t.pila,0,n.desde,n.desde)}
    const sln=Math.min(dc,nov.SLN?.dias||0);
    return {e,sal,dias:dc,diasCot:dc-sln,ibc:Math.round(sal/30*(dc-sln)),ing:e.fechaIngreso>=ini?e.fechaIngreso:"",ret:e.fechaRetiro&&e.fechaRetiro<=fin?e.fechaRetiro:"",nov};
  });
}

/* ---------- documentos: plantillas con campos entre llaves ---------- */
const CAMPOS=[["nombre","Nombre completo del empleado"],["tipo_documento","Tipo de documento"],["documento","Número de documento"],["lugar_expedicion","Lugar de expedición"],["direccion","Dirección del empleado"],["telefono","Teléfono del empleado"],["correo","Correo del empleado"],
  ["cargo","Cargo"],["salario","Salario en números"],["salario_letras","Salario en letras"],["tipo_contrato","Tipo de contrato"],["fecha_ingreso","Fecha de ingreso"],["fecha_fin","Fecha de terminación pactada"],["duracion","Duración del contrato"],["obra","Obra o labor contratada"],["lugar_trabajo","Lugar de trabajo"],["jornada","Jornada"],["periodo_pago","Forma de pago"],["vinculo","Frase «labora desde…»"],
  ["empresa","Razón social de la empresa"],["nit_empresa","NIT de la empresa"],["direccion_empresa","Dirección de la empresa"],["ciudad","Ciudad"],["telefono_empresa","Teléfono de la empresa"],["representante","Representante legal"],["documento_representante","Cédula del representante"],["firmante","Quien firma"],["cargo_firmante","Cargo de quien firma"],
  ["fecha_hoy","Fecha de hoy"],["fecha_desde","Desde (suspensión)"],["fecha_hasta","Hasta (suspensión)"],["dias","Días de suspensión"],["fecha_reintegro","Fecha de reintegro"],["fecha_retiro","Último día de trabajo"],["motivo","Causa o motivo"],["dirigido_a","A quién va dirigido"]];
const CAMPO_LBL=Object.fromEntries(CAMPOS);
const MOTIVOS=["Renuncia voluntaria presentada por el trabajador.","Mutuo acuerdo entre las partes.","Vencimiento del término pactado en el contrato a término fijo.","Terminación de la obra o labor contratada.","Terminación durante el período de prueba.","Decisión unilateral del empleador sin justa causa, con pago de la indemnización del artículo 64 del Código Sustantivo del Trabajo.","Justa causa, de acuerdo con el artículo 62 del Código Sustantivo del Trabajo: (describe aquí los hechos)."];
const ENCAB=`{{ciudad}}, {{fecha_hoy}}

Señor(a)
**{{nombre}}**
{{tipo_documento}} {{documento}}
Cargo: {{cargo}}`;
const contrato=(titulo,duracionTxt,pruebaTxt)=>`# ${titulo}

Entre los suscritos, **{{empresa}}**, identificada con NIT {{nit_empresa}}, con domicilio en {{ciudad}}, representada legalmente por {{representante}}, identificado(a) con cédula de ciudadanía número {{documento_representante}}, quien en adelante se denominará EL EMPLEADOR, y **{{nombre}}**, identificado(a) con {{tipo_documento}} número {{documento}} expedida en {{lugar_expedicion}}, con domicilio en {{direccion}}, quien en adelante se denominará EL TRABAJADOR, acuerdan celebrar el presente contrato individual de trabajo, que se rige por el Código Sustantivo del Trabajo y por las siguientes cláusulas:

**PRIMERA. Objeto.** EL TRABAJADOR se obliga a prestar sus servicios personales a EL EMPLEADOR en el cargo de **{{cargo}}**, cumpliendo las funciones propias del cargo, las que le asigne EL EMPLEADOR y las que se deriven de la naturaleza de la labor, con la dedicación, el cuidado y la diligencia debidos.

**SEGUNDA. Salario.** EL EMPLEADOR pagará a EL TRABAJADOR un salario mensual de \${{salario}} ({{salario_letras}}), {{periodo_pago}}. Dentro de este pago se entiende incluida la remuneración de los descansos dominicales y festivos. Cuando corresponda por ley, se pagará además el auxilio de transporte.

**TERCERA. Jornada.** EL TRABAJADOR cumplirá {{jornada}}. El trabajo suplementario, nocturno, dominical o festivo solo se realizará con autorización previa de EL EMPLEADOR y se pagará conforme a la ley.

**CUARTA. Lugar de trabajo.** Los servicios se prestarán en {{lugar_trabajo}}, o en el lugar que EL EMPLEADOR indique, siempre que el cambio no desmejore las condiciones de EL TRABAJADOR.

**QUINTA. Duración.** ${duracionTxt}

**SEXTA. Período de prueba.** ${pruebaTxt}

**SÉPTIMA. Obligaciones.** Además de las establecidas en la ley y en el reglamento interno de trabajo, EL TRABAJADOR se obliga a guardar reserva sobre la información de EL EMPLEADOR, de sus clientes y de sus proveedores, a cuidar los bienes que se le entreguen y a cumplir las normas de seguridad y salud en el trabajo.

**OCTAVA. Terminación.** Son justas causas para dar por terminado este contrato las señaladas en los artículos 62 y 63 del Código Sustantivo del Trabajo y las previstas en el reglamento interno de trabajo. La terminación sin justa causa dará lugar a la indemnización prevista en el artículo 64 del mismo código.

**NOVENA. Descuentos.** EL TRABAJADOR autoriza a EL EMPLEADOR para descontar de su salario y de su liquidación final las sumas que le adeude por préstamos, anticipos o elementos no devueltos, dentro de los límites que fija la ley.

**DÉCIMA. Integridad.** Este documento contiene la totalidad del acuerdo entre las partes y reemplaza cualquier acuerdo verbal o escrito anterior. Toda modificación deberá constar por escrito.

Para constancia, se firma en {{ciudad}}, el {{fecha_hoy}}, en dos (2) ejemplares del mismo tenor.

[[FIRMAS]]`;
const PLANT={
  contrato_indefinido:{l:"Contrato a término indefinido",cuerpo:contrato("Contrato individual de trabajo a término indefinido",
    "El presente contrato es a término indefinido e inicia el {{fecha_ingreso}}. Permanecerá vigente mientras subsistan las causas que le dieron origen y la materia del trabajo.",
    "Las partes acuerdan un período de prueba de dos (2) meses, durante el cual cualquiera de ellas podrá dar por terminado el contrato sin previo aviso.")},
  contrato_fijo:{l:"Contrato a término fijo",cuerpo:contrato("Contrato individual de trabajo a término fijo",
    "El presente contrato es a término fijo por {{duracion}}, contado desde el {{fecha_ingreso}} hasta el {{fecha_fin}}. Si ninguna de las partes avisa por escrito a la otra su decisión de no prorrogarlo, con una antelación no inferior a treinta (30) días a la fecha de terminación, se entenderá prorrogado en los términos del artículo 46 del Código Sustantivo del Trabajo.",
    "Las partes acuerdan un período de prueba que no excederá la quinta parte del término inicialmente pactado, sin superar dos (2) meses, durante el cual cualquiera de ellas podrá dar por terminado el contrato sin previo aviso.")},
  contrato_obra_labor:{l:"Contrato por obra o labor",cuerpo:contrato("Contrato individual de trabajo por duración de la obra o labor",
    "El presente contrato se celebra por el tiempo que dure la realización de la siguiente obra o labor: {{obra}}. Inicia el {{fecha_ingreso}} y terminará cuando la obra o labor contratada finalice, lo cual EL EMPLEADOR comunicará por escrito a EL TRABAJADOR.",
    "Las partes acuerdan un período de prueba de dos (2) meses, durante el cual cualquiera de ellas podrá dar por terminado el contrato sin previo aviso.")},
  certificado:{l:"Certificado laboral",cuerpo:`# Certificación laboral

**{{empresa}}**, identificada con NIT {{nit_empresa}},

## CERTIFICA:

Que **{{nombre}}**, identificado(a) con {{tipo_documento}} número {{documento}}, {{vinculo}}, mediante contrato de trabajo {{tipo_contrato}}, desempeñando el cargo de **{{cargo}}**, con un salario mensual de \${{salario}} ({{salario_letras}}).

La presente certificación se expide a solicitud del interesado(a), con destino a: {{dirigido_a}}.

Dada en {{ciudad}}, el {{fecha_hoy}}.

[[FIRMA_EMPRESA]]`},
  suspension:{l:"Suspensión del contrato",cuerpo:`${ENCAB}

**Asunto: Notificación de suspensión del contrato de trabajo**

Cordial saludo.

Por medio de la presente le informamos que su contrato de trabajo queda suspendido desde el {{fecha_desde}} hasta el {{fecha_hasta}}, es decir, por {{dias}} días calendario, por la siguiente causa:

{{motivo}}

Lo anterior, de conformidad con los artículos 51 y 53 del Código Sustantivo del Trabajo. Durante la suspensión se interrumpe para usted la obligación de prestar el servicio y para la empresa la de pagar los salarios de ese lapso. Usted deberá reintegrarse a sus labores el {{fecha_reintegro}}.

Si la suspensión corresponde a una sanción disciplinaria, se impone después de agotar el procedimiento previsto en el reglamento interno de trabajo, en el cual usted tuvo la oportunidad de ser escuchado(a) y de presentar sus descargos.

Atentamente,

[[FIRMA_EMPRESA]]

[[RECIBIDO]]`},
  finalizacion:{l:"Notificación de finalización (preaviso)",cuerpo:`${ENCAB}

**Asunto: Preaviso de terminación del contrato de trabajo a término fijo**

Cordial saludo.

De conformidad con el artículo 46 del Código Sustantivo del Trabajo, le informamos, con una anticipación no inferior a treinta (30) días, que la empresa ha decidido no prorrogar su contrato de trabajo a término fijo, el cual termina el {{fecha_fin}}.

En la fecha de terminación se le pagarán los salarios y las prestaciones sociales a que tenga derecho. Le pedimos coordinar con el área de recursos humanos la entrega de su cargo, el examen médico de egreso y la firma de la liquidación.

Agradecemos los servicios prestados a la empresa.

Atentamente,

[[FIRMA_EMPRESA]]

[[RECIBIDO]]`},
  terminacion:{l:"Terminación de contrato",cuerpo:`${ENCAB}

**Asunto: Terminación del contrato de trabajo**

Cordial saludo.

Por medio de la presente le comunicamos que su contrato de trabajo termina el {{fecha_retiro}}, por la siguiente causa:

{{motivo}}

La liquidación de salarios y prestaciones sociales se pagará conforme a la ley. Junto con ella le entregaremos la constancia del pago de sus aportes a seguridad social y parafiscales de los últimos tres (3) meses, como lo exige el parágrafo 1 del artículo 65 del Código Sustantivo del Trabajo, y la orden para el examen médico de egreso.

Le solicitamos hacer entrega formal de su cargo y de los elementos de trabajo que tenga a su cuidado.

Atentamente,

[[FIRMA_EMPRESA]]

[[RECIBIDO]]`},
  autorizacion:{l:"Autorización de datos personales",cuerpo:`# Autorización para el tratamiento de datos personales

Yo, **{{nombre}}**, identificado(a) con {{tipo_documento}} número {{documento}}, autorizo de manera previa, expresa e informada a **{{empresa}}**, NIT {{nit_empresa}}, para recolectar, almacenar, usar, circular y suprimir mis datos personales, incluidos los datos sensibles relacionados con mi salud y los datos de mi grupo familiar, con las siguientes finalidades: administrar la relación laboral, pagar la nómina, realizar la afiliación y el pago de aportes al sistema de seguridad social y parafiscales, cumplir las obligaciones legales del empleador y expedir certificaciones.

Declaro que se me informó que, de acuerdo con la Ley 1581 de 2012 y sus decretos reglamentarios, tengo derecho a conocer, actualizar, rectificar y suprimir mis datos, a solicitar prueba de esta autorización y a revocarla cuando proceda, mediante comunicación dirigida a {{empresa}} en {{direccion_empresa}}. También se me informó que no estoy obligado(a) a autorizar el tratamiento de datos sensibles.

Firmo en {{ciudad}}, el {{fecha_hoy}}.

[[FIRMA_TRABAJADOR]]`},
};
const DOCS={
  contrato:{l:"Contrato de trabajo",d:"Se arma solo según el tipo de contrato del empleado: indefinido, fijo u obra o labor.",plant:e=>"contrato_"+(PLANT["contrato_"+e?.tipoContrato]?e.tipoContrato:"indefinido"),campos:[]},
  certificado:{l:"Certificado laboral",d:"Constancia de cargo, fecha de ingreso, tipo de contrato y salario.",plant:()=>"certificado",campos:[["dirigido_a","A quién va dirigido","text","A quien interese"]]},
  suspension:{l:"Suspensión del contrato",d:"Notifica al empleado las fechas y la causa de la suspensión.",plant:()=>"suspension",campos:[["fecha_desde","Suspendido desde","date"],["fecha_hasta","Suspendido hasta","date"],["motivo","Causa de la suspensión","textarea"]]},
  finalizacion:{l:"Notificación de finalización",d:"Preaviso de que no se renueva un contrato a término fijo. Debe entregarse mínimo 30 días antes.",plant:()=>"finalizacion",campos:[["fecha_fin","Fecha en que termina el contrato","date"]]},
  terminacion:{l:"Terminación de contrato",d:"Carta que informa la terminación, la causa y los pasos de la liquidación.",plant:()=>"terminacion",campos:[["fecha_retiro","Último día de trabajo","date"],["motivo","Causa de la terminación","motivo"]]},
  autorizacion:{l:"Autorización de datos personales",d:"Consentimiento del empleado para el manejo de sus datos (Ley 1581 de 2012).",plant:()=>"autorizacion",campos:[]},
};
const textoPlantilla=id=>RH.plantillas[id]?.cuerpo||PLANT[id]?.cuerpo||"";
const nombreEmpresa=()=>RH.empresa.razonSocial||window.TES?.empresas?.find(x=>x.id===TES.empresa)?.nombre||"";
function valores(e,x){
  const E=RH.empresa||{},hf=hoy(),v={empresa:nombreEmpresa(),nit_empresa:E.nit||"",direccion_empresa:E.direccion||"",ciudad:E.ciudad||"",telefono_empresa:E.telefono||"",
    representante:E.representante||"",documento_representante:E.documentoRep||"",firmante:E.firmante||E.representante||"",cargo_firmante:E.cargoFirmante||"Representante legal",fecha_hoy:fechaLarga(hf)};
  if(e){const sal=salarioEn(e,e.fechaRetiro&&e.fechaRetiro<hf?e.fechaRetiro:hf);
    Object.assign(v,{nombre:nombreEmp(e),tipo_documento:TIPOS_DOC[e.tipoDoc]||e.tipoDoc||"",documento:e.documento||"",lugar_expedicion:e.expedicion||"",direccion:e.direccion||"",telefono:e.telefono||"",correo:e.correo||"",
      cargo:e.cargo||"",salario:numF.format(sal),salario_letras:sal?pesosLetras(sal):"",tipo_contrato:CONTRATO_TXT[e.tipoContrato]||"",fecha_ingreso:fechaLarga(e.fechaIngreso),fecha_fin:fechaLarga(e.fechaFin),duracion:duracion(e.fechaIngreso,e.fechaFin),obra:e.obra||"",
      lugar_trabajo:e.lugarTrabajo||[E.direccion,E.ciudad].filter(Boolean).join(", "),jornada:e.jornada||"la jornada máxima legal, en el horario que fije EL EMPLEADOR",periodo_pago:"pagadero por quincenas vencidas",
      fecha_retiro:fechaLarga(e.fechaRetiro),motivo:e.motivoRetiro||"",
      vinculo:e.fechaRetiro&&e.fechaRetiro<hf?`laboró en esta empresa desde el ${fechaLarga(e.fechaIngreso)} hasta el ${fechaLarga(e.fechaRetiro)}`:`labora en esta empresa desde el ${fechaLarga(e.fechaIngreso)}`})}
  for(const [k,val] of Object.entries(x||{})){if(val==null||val==="")continue;v[k]=/^\d{4}-\d{2}-\d{2}$/.test(val)?fechaLarga(val):val}
  if(x?.fecha_desde&&x?.fecha_hasta&&x.fecha_hasta>=x.fecha_desde){const d=diasEntre(x.fecha_desde,x.fecha_hasta)+1;v.dias=`${letras(d)} (${d})`;v.fecha_reintegro=fechaLarga(addDias(x.fecha_hasta,1))}
  return v;
}
const escH=s=>String(s??"").replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
function cuerpoDoc(texto,v){
  const falt=new Set();
  const linea=s=>escH(s).replace(/\*\*(.+?)\*\*/g,"<strong>$1</strong>").replace(/\{\{\s*([a-z_]+)\s*\}\}/g,(m,k)=>{const x=v[k];if(x==null||x===""){falt.add(k);return `<mark>[${escH(CAMPO_LBL[k]||k)}]</mark>`}return escH(x)});
  const firma=(n,...ls)=>`<td width="50%"><br><br><br>______________________________<br><strong>${escH(n||"")}</strong>${ls.filter(Boolean).map(l=>"<br>"+escH(l)).join("")}</td>`;
  const tabla=celdas=>`<table class="firmas" width="100%"><tr>${celdas}</tr></table>`;
  const idTrab=[v.tipo_documento,v.documento].filter(Boolean).join(" ");
  const html=String(texto||"").replace(/\r/g,"").split(/\n\s*\n/).map(b=>b.trim()).filter(Boolean).map(b=>{
    if(b==="[[FIRMAS]]")return tabla(firma(v.firmante,v.cargo_firmante,v.empresa,"EL EMPLEADOR")+firma(v.nombre,idTrab,"EL TRABAJADOR"));
    if(b==="[[FIRMA_EMPRESA]]")return tabla(firma(v.firmante,v.cargo_firmante,v.empresa)+"<td></td>");
    if(b==="[[FIRMA_TRABAJADOR]]")return tabla(firma(v.nombre,idTrab)+"<td></td>");
    if(b==="[[RECIBIDO]]")return tabla(firma(v.nombre,idTrab,"Recibí la presente comunicación","Fecha: ____________________")+"<td></td>");
    if(b.startsWith("## "))return `<h2>${linea(b.slice(3))}</h2>`;
    if(b.startsWith("# "))return `<h1>${linea(b.slice(2))}</h1>`;
    return `<p>${b.split("\n").map(linea).join("<br>")}</p>`}).join("");
  return {html,faltan:[...falt]};
}
const membrete=v=>RH.empresa.membrete===false?"":`<div class="membrete"><strong>${escH(v.empresa)}</strong>${v.nit_empresa?`<br>NIT ${escH(v.nit_empresa)}`:""}${v.direccion_empresa?`<br>${escH([v.direccion_empresa,v.ciudad].filter(Boolean).join(" · "))}`:""}${v.telefono_empresa?`<br>Tel. ${escH(v.telefono_empresa)}`:""}</div>`;
const CSS_DOC=`.hoja{font:11.5pt/1.55 Georgia,"Times New Roman",serif;color:#111;text-align:justify}
.hoja h1{font:bold 12.5pt/1.3 Georgia,"Times New Roman",serif;text-align:center;margin:0 0 16pt;text-transform:uppercase}
.hoja h2{font:bold 11.5pt Georgia,"Times New Roman",serif;text-align:center;margin:12pt 0 8pt}
.hoja p{margin:0 0 9pt}
.hoja .membrete{text-align:center;font-size:9.5pt;line-height:1.35;border-bottom:1px solid #999;padding-bottom:8pt;margin-bottom:18pt}
.hoja .membrete strong{font-size:12.5pt}
.hoja table{border-collapse:collapse;width:100%;font-size:10pt;margin:0 0 10pt}
.hoja td,.hoja th{border:1px solid #999;padding:4pt 6pt;text-align:left;vertical-align:top;background:none;color:#111;text-transform:none;letter-spacing:0;font-size:10pt;font-family:inherit;white-space:normal}
.hoja th{background:#eee}
.hoja td.v,.hoja th.v{text-align:right}
.hoja table.firmas{margin-top:14pt;page-break-inside:avoid}
.hoja table.firmas td{border:0;padding:0 14pt 0 0;font-size:10pt;line-height:1.35;text-align:left}
.hoja mark{background:#ffe08a;color:#111}
.salto{page-break-after:always;break-after:page;height:0}`;
document.head.append(h("style",null,CSS_DOC));
function armarDoc(tipo,e,extra){const D=DOCS[tipo],pid=D.plant(e),v=valores(e,extra),c=cuerpoDoc(textoPlantilla(pid),v);
  return {tipo,pid,titulo:D.l,html:`<div class="hoja">${membrete(v)}${c.html}</div>`,faltan:c.faltan}}
function paginaHTML(titulo,cuerpo,word){return `<!doctype html><html lang="es"${word?' xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word"':""}><head><meta charset="utf-8"><title>${escH(titulo)}</title><style>@page{size:letter;margin:2.2cm 2.4cm}body{margin:0}${CSS_DOC}</style></head><body>${cuerpo}</body></html>`}
function imprimir(titulo,partes){
  const html=paginaHTML(titulo,partes.join('<div class="salto"></div>'));
  const w=window.open("","_blank");
  if(w){w.document.open();w.document.write(html);w.document.close();setTimeout(()=>{try{w.focus();w.print()}catch(e){}},500);return}
  const f=h("iframe",{style:"position:fixed;right:0;bottom:0;width:0;height:0;border:0","aria-hidden":"true"});document.body.append(f);
  f.contentDocument.open();f.contentDocument.write(html);f.contentDocument.close();setTimeout(()=>{try{f.contentWindow.focus();f.contentWindow.print()}catch(e){toast("Permite las ventanas emergentes para imprimir.")}setTimeout(()=>f.remove(),60000)},500);
}
const descargarWord=(nombre,partes)=>guardarArchivo(nombre+".doc",new Blob(["﻿"+paginaHTML(nombre,partes.join('<br clear="all" style="page-break-before:always">'),true)],{type:"application/msword"}));
async function registrarDoc(d,e,extra){
  try{await S.db.collection("rh_documentos").add({empleadoId:e?.id||null,empleado:nombreEmp(e),tipo:d.tipo,plantilla:d.pid,titulo:d.titulo,campos:extra||{},fecha:hoy(),mes:hoy().slice(0,7),createdBy:S.uid,createdAt:nowISO()});
    registrarActividad("Generó documento",`${d.titulo} · ${nombreEmp(e)}`)}catch(err){console.warn("doc",err)}
}
const avisoFaltan=faltan=>faltan.length?h("div",{class:"status warn"},"Faltan datos (marcados en amarillo): "+uniq(faltan.map(k=>CAMPO_LBL[k]||k)).join(", ")+". Complétalos en la ficha del empleado o en Ajustes → Datos de la empresa."):null;
const archivoDoc=(d,e)=>`${d.titulo} - ${nombreEmp(e)}`.replace(/[\\/:*?"<>|]/g,"");

/* Generar un documento: elegir empleado, completar lo que pida y ver la vista previa */
function generador(tipo,opt){
  opt=opt||{};const D=DOCS[tipo];const st={emp:opt.empleadoId||"",x:{...(opt.campos||{})}};
  const prev=h("div",{class:"papel"}),av=h("div",{class:"stack"}),camposBox=h("div",{class:"cols3"});let registrado=false,actual=null;
  const E=()=>empPor(st.emp)||(opt.emp&&opt.emp.id===st.emp?opt.emp:null);
  const porDefecto=()=>{const e=E();for(const [k,,,def] of D.campos)if(st.x[k]==null&&def)st.x[k]=def;
    if(!e)return;if(tipo==="finalizacion"&&!st.x.fecha_fin)st.x.fecha_fin=e.fechaFin||"";
    if(tipo==="terminacion"){if(!st.x.fecha_retiro)st.x.fecha_retiro=e.fechaRetiro||hoy();if(!st.x.motivo)st.x.motivo=e.motivoRetiro||""}};
  const pintar=()=>{const e=E();registrado=false;
    if(!e){actual=null;prev.replaceChildren(tablaVacia("Elige el empleado para ver el documento."));av.replaceChildren();return}
    actual=armarDoc(tipo,e,st.x);prev.innerHTML=actual.html;
    const notas=[avisoFaltan(actual.faltan)];
    if(tipo==="finalizacion"){if(e.tipoContrato!=="fijo")notas.push(h("div",{class:"status warn"},"Este empleado no tiene contrato a término fijo. El preaviso solo aplica para contratos a término fijo."));
      else if(st.x.fecha_fin&&diasEntre(hoy(),st.x.fecha_fin)<30)notas.push(h("div",{class:"status bad"},`Faltan ${Math.max(0,diasEntre(hoy(),st.x.fecha_fin))} días para la terminación. El preaviso debe entregarse por escrito al menos 30 días antes; si no, el contrato se prorroga.`))}
    if(tipo==="terminacion"&&/justa causa, de acuerdo/i.test(st.x.motivo||""))notas.push(h("div",{class:"status info"},"En una terminación con justa causa, describe en la causa los hechos concretos y la fecha en que ocurrieron. Revisa el caso con tu abogado laboral antes de entregar la carta."));
    av.replaceChildren(...notas.filter(Boolean))};
  const pintarCampos=()=>camposBox.replaceChildren(...D.campos.map(([k,l,t])=>{
    const on=ev=>{st.x[k]=ev.target.value;pintar()};
    const el=t==="textarea"?h("textarea",{id:"gd-"+k,oninput:on},st.x[k]||""):t==="motivo"?h("textarea",{id:"gd-"+k,oninput:on},st.x[k]||""):h("input",{id:"gd-"+k,type:t,value:st.x[k]||"",oninput:on});
    const extra=t==="motivo"?h("select",{id:"gd-motivos","aria-label":"Causas frecuentes",onchange:ev=>{if(!ev.target.value)return;st.x[k]=ev.target.value;el.value=ev.target.value;pintar()}},h("option",{value:""},"Elegir una causa frecuente…"),MOTIVOS.map(m=>h("option",{value:m},m.length>70?m.slice(0,70)+"…":m))):null;
    return h("div",{class:"stack",style:"gap:6px;"+(t==="textarea"||t==="motivo"?"grid-column:1/-1":"")},campo(l,el,true),extra)}));
  const usar=async fn=>{if(!actual){toast("Elige el empleado.");return}const e=E();if(!registrado){registrado=true;registrarDoc(actual,e,st.x)}fn(actual,e)};
  porDefecto();pintarCampos();pintar();
  abrirModal(D.l,h("div",{class:"stack"},
    h("p",{class:"small muted",style:"margin:0"},D.d),
    campo("Empleado",selEmp(st.emp,false,v=>{st.emp=v;porDefecto();pintarCampos();pintar()},"gd-emp"),true),
    camposBox,av,prev,
    h("div",{class:"row"},h("button",{class:"pri",onclick:()=>usar((d,e)=>imprimir(archivoDoc(d,e),[d.html]))},"Imprimir o guardar en PDF"),
      h("button",{onclick:()=>usar((d,e)=>descargarWord(archivoDoc(d,e),[d.html]))},"Descargar en Word"),
      h("span",{class:"small muted"},"El texto se cambia en Ajustes → Plantillas."))),{ancho:900});
}
/* Paquete de documentos de ingreso: contrato + autorización de datos */
function paquete(e,tipos){
  const ds=tipos.map(t=>armarDoc(t,e,{}));let registrado=false;
  const reg=()=>{if(registrado)return;registrado=true;ds.forEach(d=>registrarDoc(d,e,{}))};
  abrirModal("Documentos de ingreso de "+nombreEmp(e),h("div",{class:"stack"},
    h("div",{class:"status ok"},"El empleado quedó guardado. Estos son sus documentos de ingreso, listos para imprimir y firmar."),
    avisoFaltan(ds.flatMap(d=>d.faltan)),
    ds.map(d=>{const p=h("div",{class:"papel",style:"max-height:36vh"});p.innerHTML=d.html;return h("div",{class:"stack",style:"gap:6px"},h("h3",null,d.titulo),p)}),
    h("div",{class:"row"},h("button",{class:"pri",onclick:()=>{reg();imprimir("Documentos de ingreso - "+nombreEmp(e),ds.map(d=>d.html))}},"Imprimir todo o guardar en PDF"),
      h("button",{onclick:()=>{reg();descargarWord("Documentos de ingreso - "+nombreEmp(e).replace(/[\\/:*?"<>|]/g,""),ds.map(d=>d.html))}},"Descargar en Word"))),{ancho:900});
}

/* ---------- empleados ---------- */
function formEmpleado(e){
  const I={},nuevo=!e,p=P();const msg=h("div",{class:"status bad",hidden:true});let cerrar=null;
  const F=(k,l,o)=>{o=o||{};const val=String(e?.[k]??o.def??"");
    const el=o.sel?h("select",{id:"em-"+k},o.sel.map(([v,t])=>h("option",{value:v,selected:val===v?true:null},t))):h("input",{id:"em-"+k,type:o.type||"text",value:val,placeholder:o.ph||null,inputmode:o.num?"decimal":null,autocomplete:"off"});
    I[k]=el;return campo(l,el,o.req,o.ayuda)};
  const cFin=F("fechaFin","Fecha en que termina el contrato",{type:"date",req:1}),cObra=F("obra","Obra o labor contratada",{req:1,ph:"Ej. Instalación de redes del proyecto Torre Norte"});
  const tipoC=F("tipoContrato","Tipo de contrato",{req:1,def:"indefinido",sel:Object.entries(CONTRATO_LBL)});
  const ver=()=>{cFin.hidden=I.tipoContrato.value!=="fijo";cObra.hidden=I.tipoContrato.value!=="obra_labor";cObra.style.gridColumn="1/-1"};
  I.tipoContrato.addEventListener("change",ver);ver();
  const gen=nuevo?h("input",{type:"checkbox",id:"em-generar",checked:true}):null;
  const btn=h("button",{class:"pri",type:"submit"},nuevo?"Guardar empleado":"Guardar cambios");
  const sec=(t,...c)=>h("div",{class:"stack sec-form"},h("h3",null,t),h("div",{class:"cols3"},c));
  const form=h("form",{class:"stack",onsubmit:async ev=>{ev.preventDefault();
    const d={};for(const k of Object.keys(I))d[k]=I[k].value.trim();d.salario=Math.abs(parseMonto(d.salario)||0);
    const mal=t=>{msg.hidden=false;msg.textContent=t;msg.scrollIntoView({block:"nearest"})};
    const req=[["primerNombre","el primer nombre"],["primerApellido","el primer apellido"],["documento","el número de documento"],["cargo","el cargo"],["fechaIngreso","la fecha de ingreso"]];
    for(const [k,l] of req)if(!d[k]){I[k].focus();return mal("Falta "+l+".")}
    if(!d.salario)return mal("Escribe el salario mensual.");
    if(d.tipoContrato==="fijo"){if(!d.fechaFin)return mal("Indica la fecha en que termina el contrato a término fijo.");if(d.fechaFin<=d.fechaIngreso)return mal("La fecha de terminación debe ser posterior a la fecha de ingreso.")}else d.fechaFin="";
    if(d.tipoContrato==="obra_labor"){if(!d.obra)return mal("Describe la obra o labor contratada.")}else d.obra="";
    if(RH.emp.some(x=>x.id!==e?.id&&soloNum(x.documento)===soloNum(d.documento)))return mal("Ya existe un empleado con ese número de documento.");
    btn.disabled=true;btn.textContent="Guardando…";const at=nowISO();
    try{
      if(nuevo){const id=rid(),doc={...d,fechaRetiro:"",motivoRetiro:"",createdBy:S.uid,createdAt:at,updatedBy:S.uid,updatedAt:at,historial:[{at,by:S.uid,accion:"Creado"}]};
        await S.db.doc("rh_empleados/"+id).set(doc);registrarActividad("Registró empleado",`${nombreEmp(d)} · ${d.cargo}`);
        cerrar();toast("Empleado guardado.");if(gen.checked)paquete({id,...doc},["contrato","autorizacion"])}
      else{const cambios=Object.keys(d).filter(k=>String(d[k]??"")!==String(e[k]??"")).map(k=>({campo:k,antes:e[k]??"",despues:d[k]}));
        if(cambios.length){await S.db.doc("rh_empleados/"+e.id).update({...d,updatedBy:S.uid,updatedAt:at,historial:[...(e.historial||[]),{at,by:S.uid,accion:"Modificado",cambios}].slice(-60)});
          registrarActividad("Modificó empleado",`${nombreEmp(d)} · ${cambios.map(c=>c.campo).join(", ")}`)}
        cerrar();toast("Cambios guardados.")}
    }catch(err){mal("No se pudo guardar: "+(err?.message||"revisa tu acceso e intenta de nuevo")+".");btn.disabled=false;btn.textContent=nuevo?"Guardar empleado":"Guardar cambios"}
  }},
    h("p",{class:"small muted",style:"margin:0"},"Los campos con * son obligatorios. Lo demás se puede completar después, pero entre más datos tenga la ficha, más completos salen los documentos."),
    h("div",{class:"stack"},h("h3",null,"Datos personales"),h("div",{class:"cols3"},
      F("primerNombre","Primer nombre",{req:1}),F("segundoNombre","Segundo nombre"),F("primerApellido","Primer apellido",{req:1}),F("segundoApellido","Segundo apellido"),
      F("tipoDoc","Tipo de documento",{req:1,def:"CC",sel:[["CC","Cédula de ciudadanía"],["CE","Cédula de extranjería"],["TI","Tarjeta de identidad"],["PA","Pasaporte"],["PT","Permiso por protección temporal"]]}),
      F("documento","Número de documento",{req:1}),F("expedicion","Lugar de expedición",{ph:"Ej. Medellín"}),F("fechaNacimiento","Fecha de nacimiento",{type:"date"}),
      F("direccion","Dirección"),F("ciudadRes","Ciudad de residencia"),F("telefono","Teléfono"),F("correo","Correo",{type:"email"}))),
    sec("Contrato",F("cargo","Cargo",{req:1}),F("area","Área"),tipoC,F("fechaIngreso","Fecha de ingreso",{type:"date",req:1,def:hoy()}),cFin,
      F("salario","Salario mensual",{req:1,num:1,ph:String(p.smmlv),ayuda:nuevo?`Salario mínimo vigente: ${fmt(p.smmlv)}`:"Para un aumento usa Novedades → Cambio de salario: así queda la fecha desde la que rige."}),
      F("lugarTrabajo","Lugar de trabajo",{ph:"Si se deja vacío, la sede de la empresa"}),F("jornada","Jornada u horario",{ph:"Opcional. Ej. lunes a viernes de 8 a. m. a 5 p. m."}),cObra),
    sec("Seguridad social",F("eps","EPS (salud)"),F("afp","Fondo de pensiones"),F("cesantias","Fondo de cesantías"),F("arl","ARL"),
      F("riesgo","Clase de riesgo ARL",{def:"1",sel:[["1","I · mínimo (oficina)"],["2","II · bajo"],["3","III · medio"],["4","IV · alto"],["5","V · máximo"]]}),F("caja","Caja de compensación")),
    sec("Cuenta para el pago",F("banco","Banco"),F("tipoCuenta","Tipo de cuenta",{def:"Ahorros",sel:[["Ahorros","Ahorros"],["Corriente","Corriente"],["Billetera","Billetera digital (Nequi, Daviplata)"],["Efectivo","Pago en efectivo"]]}),F("numeroCuenta","Número de cuenta")),
    nuevo?h("label",{class:"row",for:"em-generar",style:"gap:8px"},gen,h("span",null,h("strong",null,"Generar los documentos de ingreso al guardar")," (contrato y autorización de datos)")):null,
    msg,h("div",{class:"row"},btn));
  cerrar=abrirModal(nuevo?"Nuevo empleado":"Editar a "+nombreEmp(e),form,{ancho:900});
}
function retirar(e){
  const I={};const msg=h("div",{class:"status bad",hidden:true});let cerrar=null;
  const gen=h("input",{type:"checkbox",id:"rt-gen",checked:true});const btn=h("button",{class:"pri",type:"submit"},"Registrar retiro");
  cerrar=abrirModal("Retirar a "+nombreEmp(e),h("form",{class:"stack",onsubmit:async ev=>{ev.preventDefault();
    const f=I.fecha.value,m=I.motivo.value.trim();if(!f||f<e.fechaIngreso){msg.hidden=false;msg.textContent="Indica el último día de trabajo.";return}if(!m){msg.hidden=false;msg.textContent="Indica la causa del retiro.";return}
    btn.disabled=true;const at=nowISO();
    try{await S.db.doc("rh_empleados/"+e.id).update({fechaRetiro:f,motivoRetiro:m,updatedBy:S.uid,updatedAt:at,historial:[...(e.historial||[]),{at,by:S.uid,accion:"Retiro registrado: "+fmtF(f)}].slice(-60)});
      registrarActividad("Registró retiro de empleado",`${nombreEmp(e)} · ${fmtF(f)}`);cerrar();toast("Retiro registrado.");
      if(gen.checked)generador("terminacion",{empleadoId:e.id,emp:{...e,fechaRetiro:f,motivoRetiro:m},campos:{fecha_retiro:f,motivo:m}})}
    catch(err){msg.hidden=false;msg.textContent="No se pudo guardar.";btn.disabled=false}}},
    h("p",{class:"small muted",style:"margin:0"},"El empleado deja de aparecer en la nómina después de su último día y sale con novedad de retiro en la planilla de seguridad social."),
    campo("Último día de trabajo",I.fecha=h("input",{id:"rt-fecha",type:"date",value:hoy()}),true),
    campo("Causa del retiro",I.motivo=h("textarea",{id:"rt-motivo"}),true),
    h("select",{id:"rt-motivos","aria-label":"Causas frecuentes",onchange:ev=>{if(ev.target.value)I.motivo.value=ev.target.value}},h("option",{value:""},"Elegir una causa frecuente…"),MOTIVOS.map(m=>h("option",{value:m},m.length>70?m.slice(0,70)+"…":m))),
    h("label",{class:"row",for:"rt-gen",style:"gap:8px"},gen,"Generar la carta de terminación del contrato"),msg,h("div",{class:"row"},btn)),{ancho:560});
}
function fichaEmpleado(e){
  const est=estadoEmp(e),sal=salarioEn(e,hoy());let cerrar=null;
  const dd=(l,v)=>v?[h("dt",null,l),h("dd",null,v)]:null;
  const novs=RH.nov.filter(n=>n.empleadoId===e.id&&!n.anulado).sort((a,b)=>b.desde.localeCompare(a.desde)).slice(0,8);
  const docs=RH.docs.filter(d=>d.empleadoId===e.id).sort((a,b)=>String(b.createdAt).localeCompare(String(a.createdAt))).slice(0,8);
  const selDoc=h("select",{id:"fe-doc","aria-label":"Generar documento",style:"max-width:240px",onchange:ev=>{const t=ev.target.value;ev.target.value="";if(t){cerrar();generador(t,{empleadoId:e.id})}}},h("option",{value:""},"Generar documento…"),Object.entries(DOCS).map(([k,d])=>h("option",{value:k},d.l)));
  cerrar=abrirModal(nombreEmp(e),h("div",{class:"stack"},
    h("div",{class:"row"},est.chip,h("span",{class:"chip neu"},CONTRATO_LBL[e.tipoContrato]||"Contrato"),e.tipoContrato==="fijo"&&e.fechaFin?h("span",{class:"chip info"},"Vence el "+fmtF(e.fechaFin)):null),
    h("div",{class:"row"},h("button",{class:"pri",onclick:()=>{cerrar();formEmpleado(e)}},"Editar ficha"),h("button",{onclick:()=>{cerrar();modalNovedad(e.id)}},"Registrar novedad"),selDoc,
      est.k==="activo"&&!e.fechaRetiro?h("button",{class:"danger",onclick:()=>{cerrar();retirar(e)}},"Retirar"):
      h("button",{onclick:ev=>confirmarEnPagina(ev.currentTarget,"Confirmar reintegro",async()=>{try{await S.db.doc("rh_empleados/"+e.id).update({fechaRetiro:"",motivoRetiro:"",updatedBy:S.uid,updatedAt:nowISO()});registrarActividad("Deshizo retiro de empleado",nombreEmp(e));cerrar();toast("El empleado vuelve a estar activo.")}catch(err){toast("No se pudo guardar.")}})},"Deshacer retiro")),
    h("dl",{class:"dl"},dd("Documento",`${(e.tipoDoc||"CC")} ${e.documento}${e.expedicion?" de "+e.expedicion:""}`),dd("Cargo",[e.cargo,e.area].filter(Boolean).join(" · ")),dd("Ingreso",fmtF(e.fechaIngreso)),
      dd("Salario actual",fmt(sal)+(sal!==Number(e.salario)?` (ingresó con ${fmt(e.salario)})`:"")),dd("Retiro",e.fechaRetiro?`${fmtF(e.fechaRetiro)} · ${e.motivoRetiro||""}`:""),
      dd("Contacto",[e.telefono,e.correo,e.direccion,e.ciudadRes].filter(Boolean).join(" · ")),dd("Nacimiento",fmtF(e.fechaNacimiento)),
      dd("Seguridad social",[e.eps&&"EPS "+e.eps,e.afp&&"Pensión "+e.afp,e.cesantias&&"Cesantías "+e.cesantias,e.arl&&`ARL ${e.arl} (riesgo ${e.riesgo||1})`,e.caja&&"Caja "+e.caja].filter(Boolean).join(" · ")),
      dd("Pago",[e.banco,e.tipoCuenta,e.numeroCuenta].filter(Boolean).join(" · ")),dd("Registrado por",`${nombre(e.createdBy)} · ${fmtFH(e.createdAt)}`)),
    h("h3",null,"Últimas novedades"),novs.length?h("div",{class:"tbl"},h("table",null,h("tbody",null,novs.map(n=>{const t=novTexto(n);return h("tr",null,h("td",null,NOV[n.tipo]?.l||n.tipo),h("td",null,t.cuando),h("td",null,t.cuanto))})))):h("div",{class:"small muted"},"Sin novedades registradas."),
    h("h3",null,"Documentos generados"),docs.length?h("div",{class:"tbl"},h("table",null,h("tbody",null,docs.map(d=>h("tr",null,h("td",null,d.titulo),h("td",null,fmtFH(d.createdAt)),h("td",null,nombre(d.createdBy))))))):h("div",{class:"small muted"},"Aún no se le ha generado ningún documento.")),{ancho:820});
}
VIEWS.rh_empleados={busca:"",estado:"activo",
  mount(el){iniciar();
    this.kpis=h("div",{class:"kpis"});this.lista=h("div");this.filtros=h("div",{class:"filters"});
    const busca=h("input",{id:"rh-busca",type:"search",placeholder:"Buscar por nombre, documento o cargo",style:"max-width:300px",value:this.busca,oninput:e=>{this.busca=e.target.value;this.refresh()}});
    el.append(h("div",{class:"row between"},h("div",null,h("h2",null,"Empleados"),h("div",{class:"small muted"},"El cuadro de personal. Toca un empleado para ver su ficha, registrar novedades o generar sus documentos.")),
        h("button",{class:"pri",onclick:()=>formEmpleado(null)},"Nuevo empleado")),
      this.kpis,h("div",{class:"card"},h("div",{class:"row between",style:"margin-bottom:12px"},this.filtros,busca),this.lista));
  },
  refresh(){
    const hf=hoy(),act=RH.emp.filter(e=>estadoEmp(e).k==="activo"),ret=RH.emp.length-act.length;
    const vencen=act.filter(e=>e.tipoContrato==="fijo"&&e.fechaFin&&!e.fechaRetiro&&diasEntre(hf,e.fechaFin)>=0&&diasEntre(hf,e.fechaFin)<=45);
    this.kpis.replaceChildren(kpi("Empleados activos",String(act.length),ret?`${ret} retirados`:"sin retirados"),
      kpi("Nómina mensual",fmt(act.reduce((s,e)=>s+salarioEn(e,hf),0)),"suma de salarios básicos"),
      kpi("Contratos por vencer",String(vencen.length),vencen.length?"en los próximos 45 días":"ninguno en 45 días",vencen.length>0));
    this.filtros.replaceChildren(...[["activo","Activos"],["retirado","Retirados"],["todos","Todos"]].map(([k,l])=>h("button",{class:"sm","aria-pressed":String(this.estado===k),onclick:()=>{this.estado=k;this.refresh()}},l)));
    if(!RH.ok.emp)return this.lista.replaceChildren(tablaVacia("Cargando empleados…"));
    const q=norm(this.busca);
    const l=RH.emp.filter(e=>this.estado==="todos"||estadoEmp(e).k===this.estado).filter(e=>!q||norm(`${nombreEmp(e)} ${e.documento} ${e.cargo} ${e.area}`).includes(q)).sort(porNombre);
    this.lista.replaceChildren(l.length?h("div",{class:"tbl"},h("table",null,h("thead",null,h("tr",null,h("th",null,"Empleado"),h("th",null,"Cargo"),h("th",null,"Contrato"),h("th",null,"Ingreso"),h("th",{class:"v"},"Salario"),h("th",null,"Estado"))),
      h("tbody",null,l.map(e=>h("tr",{class:"click",onclick:()=>fichaEmpleado(e)},h("td",null,h("strong",null,nombreEmp(e)),h("div",{class:"small muted"},`${e.tipoDoc||"CC"} ${e.documento}`)),h("td",null,e.cargo,e.area?h("div",{class:"small muted"},e.area):null),
        h("td",null,CONTRATO_LBL[e.tipoContrato]||"—",e.tipoContrato==="fijo"&&e.fechaFin?h("div",{class:"small muted"},"hasta "+fmtF(e.fechaFin)):null),h("td",null,fmtF(e.fechaIngreso)),h("td",{class:"v"},fmt(salarioEn(e,hf))),h("td",null,estadoEmp(e).chip))))))
      :tablaVacia(RH.emp.length?"No hay empleados con este filtro.":"Aún no hay empleados. Presiona «Nuevo empleado» para crear el primero: al guardarlo se generan su contrato y sus documentos de ingreso."));
  }
};

/* ---------- novedades ---------- */
function formNovedad(pre,alGuardar){
  const I={};const dyn=h("div",{class:"stack"}),ayuda=h("div",{class:"small muted"}),msg=h("div",{class:"status bad",hidden:true}),tras=h("div");
  const empBox=h("div");const pintarEmp=v=>{I.emp=selEmp(v,true,()=>calc(),"nv-emp");empBox.replaceChildren(I.emp)};
  const grupos=uniq(Object.values(NOV).map(t=>t.g));
  I.tipo=h("select",{id:"nv-tipo",onchange:()=>pintar()},grupos.map(g=>h("optgroup",{label:g},Object.entries(NOV).filter(([,t])=>t.g===g).map(([k,t])=>h("option",{value:k},t.l)))));
  const cuenta=h("div",{class:"small",style:"font-weight:600"});
  const calc=()=>{const t=NOV[I.tipo.value],e=empPor(I.emp.value);
    if(t.dias)cuenta.textContent=I.desde.value&&I.hasta.value&&I.hasta.value>=I.desde.value?`${diasEntre(I.desde.value,I.hasta.value)+1} días`:"";
    else if(t.horas)cuenta.textContent=e&&Number(I.horas.value)>0?`Valor calculado: ${fmt(valorHoras({horas:I.horas.value,subtipo:I.subtipo.value},salarioEn(e,I.desde.value||hoy()),P()))}`:"";
    else cuenta.textContent=parseMonto(I.valor.value)?fmt(Math.abs(parseMonto(I.valor.value))):""};
  const pintar=()=>{const k=I.tipo.value,t=NOV[k];ayuda.textContent=t.ayuda||"";
    if(t.dias)dyn.replaceChildren(h("div",{class:"row",style:"flex-wrap:nowrap"},h("div",{style:"flex:1"},campo("Desde",I.desde=h("input",{id:"nv-desde",type:"date",value:hoy(),onchange:()=>{if(!I.hasta.value||I.hasta.value<I.desde.value)I.hasta.value=I.desde.value;calc()}}),true)),
      h("div",{style:"flex:1"},campo("Hasta",I.hasta=h("input",{id:"nv-hasta",type:"date",value:hoy(),onchange:calc}),true))),cuenta);
    else if(t.horas)dyn.replaceChildren(campo("Tipo de hora",I.subtipo=h("select",{id:"nv-subtipo",onchange:calc},Object.entries(HORAS).map(([c,[l]])=>h("option",{value:c},l))),true),
      h("div",{class:"row",style:"flex-wrap:nowrap"},h("div",{style:"flex:1"},campo("Cantidad de horas",I.horas=h("input",{id:"nv-horas",inputmode:"decimal",oninput:calc}),true)),h("div",{style:"flex:1"},campo("Fecha",I.desde=h("input",{id:"nv-desde",type:"date",value:hoy(),onchange:calc}),true))),cuenta);
    else dyn.replaceChildren(h("div",{class:"row",style:"flex-wrap:nowrap"},h("div",{style:"flex:1"},campo(k==="cambio_salario"?"Nuevo salario mensual":"Valor",I.valor=h("input",{id:"nv-valor",inputmode:"decimal",placeholder:"0",oninput:calc}),true)),
      h("div",{style:"flex:1"},campo(k==="cambio_salario"?"Rige desde":"Fecha",I.desde=h("input",{id:"nv-desde",type:"date",value:hoy()}),true))),cuenta);
    calc()};
  I.detalle=h("textarea",{id:"nv-detalle",placeholder:"Opcional. Ej. número de la incapacidad, motivo del permiso"});
  const btn=h("button",{class:"pri",type:"submit"},"Guardar novedad");
  pintarEmp(pre||"");pintar();
  const form=h("form",{class:"stack",onsubmit:async ev=>{ev.preventDefault();
    const k=I.tipo.value,t=NOV[k],e=empPor(I.emp.value);const mal=x=>{msg.hidden=false;msg.textContent=x};msg.hidden=true;
    if(!e)return mal("Elige el empleado.");if(!I.desde.value)return mal("Indica la fecha.");
    const d={empleadoId:e.id,tipo:k,desde:I.desde.value,hasta:"",subtipo:"",horas:0,valor:0,detalle:I.detalle.value.trim()};
    if(t.dias){d.hasta=I.hasta.value||d.desde;if(d.hasta<d.desde)return mal("La fecha final no puede ser anterior a la inicial.")}
    else if(t.horas){d.subtipo=I.subtipo.value;d.horas=Math.abs(parseMonto(I.horas.value)||0);if(!d.horas)return mal("Escribe la cantidad de horas.")}
    else{d.valor=Math.abs(parseMonto(I.valor.value)||0);if(!d.valor)return mal("Escribe el valor.")}
    if(d.desde<e.fechaIngreso)return mal("La fecha es anterior al ingreso del empleado ("+fmtF(e.fechaIngreso)+").");
    btn.disabled=true;const at=nowISO();
    try{await S.db.collection("rh_novedades").add({...d,mes:d.desde.slice(0,7),anulado:false,createdBy:S.uid,createdAt:at});
      registrarActividad("Registró novedad",`${nombreEmp(e)} · ${t.l} · ${novTexto(d).cuanto}`);toast("Novedad guardada. Ya se tiene en cuenta en la nómina.");
      tras.replaceChildren(t.doc?h("div",{class:"status ok row between"},h("span",null,`Novedad guardada para ${nombreEmp(e)}.`),h("button",{type:"button",class:"sm pri",onclick:()=>generador(t.doc,{empleadoId:e.id,campos:{fecha_desde:d.desde,fecha_hasta:d.hasta,motivo:d.detalle}})},"Generar la carta de "+DOCS[t.doc].l.toLowerCase())):"");
      I.detalle.value="";pintar();if(alGuardar)alGuardar(d)}
    catch(err){mal("No se pudo guardar: "+(err?.message||"intenta de nuevo")+".")}
    btn.disabled=false}},
    campo("Empleado",empBox,true),campo("Qué pasó",I.tipo,true),ayuda,dyn,campo("Detalle",I.detalle),msg,h("div",{class:"row"},btn),tras);
  form.actualizarEmpleados=()=>{if(document.activeElement!==I.emp)pintarEmp(I.emp.value)};
  return form;
}
function modalNovedad(empleadoId){const cerrar=abrirModal("Registrar novedad",formNovedad(empleadoId,d=>{if(!NOV[d.tipo].doc)cerrar()}),{ancho:560})}
VIEWS.rh_novedades={filtro:"",
  mount(el){iniciar();
    this.form=formNovedad("");this.lista=h("div");this.sel=h("div");
    el.append(h("div",null,h("h2",{id:"rh-nvt"},"Novedades"),h("div",{class:"small muted"},"Todo lo que cambia el pago o el contrato de un empleado: vacaciones, incapacidades, permisos, horas extra, descuentos, aumentos. La nómina y la planilla de seguridad social las toman de aquí.")),
      h("div",{class:"grid2"},h("div",{class:"card"},h("h3",null,"Registrar novedad"),this.form),h("div",{class:"card"},h("div",{class:"row between",style:"margin-bottom:12px"},h("h3",{id:"rh-nvl",style:"margin:0"},"Novedades del mes"),this.sel),this.lista)));
  },
  refresh(){
    this.form.actualizarEmpleados();
    $("#rh-nvl").textContent="Novedades de "+mesNombre(S.mes).toLowerCase();
    this.sel.replaceChildren(h("select",{id:"nv-filtro","aria-label":"Filtrar por empleado",style:"max-width:240px",onchange:e=>{this.filtro=e.target.value;this.refresh()}},h("option",{value:""},"Todos los empleados"),RH.emp.slice().sort(porNombre).map(e=>h("option",{value:e.id,selected:this.filtro===e.id?true:null},nombreEmp(e)))));
    if(cargandoRH())return this.lista.replaceChildren(tablaVacia("Cargando…"));
    const ini=S.mes+"-01",fin=S.mes+"-31";
    const l=RH.nov.filter(n=>n.desde<=fin&&(n.hasta||n.desde)>=ini&&(!this.filtro||n.empleadoId===this.filtro)).sort((a,b)=>(b.desde+b.createdAt).localeCompare(a.desde+a.createdAt));
    this.lista.replaceChildren(l.length?h("div",{class:"tbl"},h("table",null,h("thead",null,h("tr",null,h("th",null,"Empleado"),h("th",null,"Novedad"),h("th",null,"Cuándo"),h("th",null,"Cuánto"),h("th",null,"Registró"),h("th",null,""))),
      h("tbody",null,l.map(n=>{const t=NOV[n.tipo]||{l:n.tipo},x=novTexto(n),e=empPor(n.empleadoId);
        return h("tr",{class:n.anulado?"tachado":null},h("td",null,e?nombreEmp(e):"—"),h("td",null,t.l,n.detalle?h("div",{class:"small muted"},n.detalle):null),h("td",null,x.cuando),h("td",null,x.cuanto),h("td",{class:"small"},nombre(n.createdBy),h("div",{class:"muted"},fmtFH(n.createdAt))),
          h("td",null,n.anulado?h("span",{class:"chip neu"},"Anulada"):h("div",{class:"row",style:"gap:6px;flex-wrap:nowrap"},
            t.doc&&e?h("button",{class:"sm",onclick:()=>generador(t.doc,{empleadoId:e.id,campos:{fecha_desde:n.desde,fecha_hasta:n.hasta,motivo:n.detalle}})},"Carta"):null,
            h("button",{class:"sm danger",onclick:ev=>confirmarEnPagina(ev.currentTarget,"Confirmar",async()=>{try{await S.db.doc("rh_novedades/"+n.id).update({anulado:true,anuladoPor:S.uid,anuladoEn:nowISO()});registrarActividad("Anuló novedad",`${e?nombreEmp(e):""} · ${t.l}`);toast("Novedad anulada.")}catch(err){toast("No se pudo anular.")}})},"Anular"))))}))))
      :tablaVacia("No hay novedades en este mes."));
  }
};

/* ---------- nómina quincenal ---------- */
function desprendible(L,per){
  /* mismo formato del comprobante que ya usa la empresa: ingresos y deducciones con concepto, cantidad y valor */
  const v=valores(null,{}),n0=x=>Number(x)||0,pesos=x=>x?"$"+numF.format(x):"-",otros=n0(L.desc)-n0(L.prest),hExtraV=n0(L.extras)-n0(L.vDom);
  const dia=f=>`${+f.slice(8,10)} DE ${MESES[+f.slice(5,7)-1].toUpperCase()}`;
  const ing=[["Sueldo",L.diasTrab,L.basico],["Auxilio de Transporte",L.diasTrab,L.aux],["Auxilio no constitutivo de salario",L.dias,L.noSal],["Horas extras",n0(L.hExtra),hExtraV],["Horas ordinarias dominicales",n0(L.hDom),n0(L.vDom)],
      ...(L.ausPag?[["Vacaciones y licencias pagadas",L.diasPagados,L.ausPag]]:[]),...(L.incap?[["Incapacidades",L.diasIncap,L.incap]]:[]),...(L.comis?[["Comisiones y bonificaciones",1,L.comis]]:[])],
    ded=[["Fondo de Salud",L.dias,L.salud],["Fondo de Pensión",L.dias,L.pension],["Descuento préstamo",L.dias,n0(L.prest)],...(L.fsp?[["Fondo de solidaridad pensional",L.dias,L.fsp]]:[]),...(otros>0?[["Otros descuentos",1,otros]]:[])];
  const c3=x=>x?`<td>${escH(x[0])}</td><td class="v">${x[1]}</td><td class="v">${pesos(x[2])}</td>`:"<td></td><td></td><td></td>";
  const filas=Array.from({length:Math.max(ing.length,ded.length)},(_,i)=>`<tr>${c3(ing[i])}${c3(ded[i])}</tr>`).join("");
  return `<div class="hoja" style="text-align:left"><p style="text-align:center;margin:0"><strong style="font-size:13pt">${escH(v.empresa)}</strong>${v.nit_empresa?`<br>NIT: ${escH(v.nit_empresa)}`:""}</p>
<h1 style="margin:12pt 0">Comprobante de nómina</h1>
<table><tr><td width="32%"><strong>PERIODO DE PAGO:</strong></td><td>${dia(per.desde)} AL ${dia(per.hasta)} DE ${per.hasta.slice(0,4)}</td></tr>
<tr><td><strong>NOMBRE:</strong></td><td>${escH(String(L.nombre).toUpperCase())}</td></tr>
<tr><td><strong>IDENTIFICACIÓN:</strong></td><td>${escH(L.documento)}</td></tr>
<tr><td><strong>CARGO:</strong></td><td>${escH(String(L.cargo).toUpperCase())}</td></tr>
<tr><td><strong>SALARIO BÁSICO:</strong></td><td>$${numF.format(L.salario)}</td></tr>
<tr><td><strong>TIPO DE VINCULACIÓN:</strong></td><td>NÓMINA</td></tr></table>
<table><tr><th colspan="3" style="text-align:center">INGRESOS</th><th colspan="3" style="text-align:center">DEDUCCIONES</th></tr>
<tr><th>CONCEPTO</th><th class="v">CANTIDAD</th><th class="v">VALOR</th><th>CONCEPTO</th><th class="v">CANTIDAD</th><th class="v">VALOR</th></tr>${filas}
<tr><th colspan="2">TOTAL INGRESOS</th><th class="v">${pesos(L.devengado)}</th><th colspan="2">TOTAL DEDUCCIONES</th><th class="v">${pesos(L.deducciones)}</th></tr>
<tr><th colspan="5" style="text-align:right">NETO A PAGAR</th><th class="v">$${numF.format(L.neto)}</th></tr></table>
<table class="firmas" width="100%"><tr><td width="50%"><br><br>______________________________<br><strong>${escH(L.nombre)}</strong><br>Recibí conforme</td><td></td></tr></table></div>`;
}
VIEWS.rh_nomina={q:null,
  mount(el){iniciar();if(this.q==null)this.q=+hoy().slice(8,10)<=15?1:2;
    this.qs=h("div",{class:"filters"});this.box=h("div",{class:"stack"});
    el.append(h("div",{class:"row between"},h("div",null,h("h2",{id:"rh-nt"},"Nómina quincenal"),h("div",{class:"small muted"},"Se calcula sola con los empleados y las novedades. Cambia el mes arriba y elige la quincena.")),this.qs),this.box);
  },
  refresh(){
    const per=periodo(S.mes,this.q),p=P();$("#rh-nt").textContent="Nómina · "+per.nombre;
    this.qs.replaceChildren(...[1,2].map(q=>h("button",{"aria-pressed":String(this.q===q),onclick:()=>{this.q=q;this.refresh()}},q===1?"Del 1 al 15":"Del 16 al fin de mes")));
    if(cargandoRH())return this.box.replaceChildren(tablaVacia("Cargando…"));
    const N=nominaDe(per),f=N.filas,T=k=>suma(f,k);
    if(!f.length)return this.box.replaceChildren(h("div",{class:"card"},tablaVacia(RH.emp.length?"Ningún empleado estuvo activo en esta quincena.":"Aún no hay empleados. Créalos en la pestaña Empleados y aquí aparecerá la nómina.")));
    const estado=N.cerrada?h("div",{class:"status ok row between"},h("span",null,`Quincena cerrada por ${nombre(N.doc.cerradaPor)} el ${fmtFH(N.doc.cerradaEn)}. Los cambios posteriores en empleados o novedades ya no la modifican.`,N.doc.enviada?` Enviada a Conciliación con fecha de pago ${fmtF(N.doc.enviada.fecha)}.`:""),
        esAdmin()?h("button",{class:"sm",onclick:ev=>confirmarEnPagina(ev.currentTarget,"Confirmar",async()=>{try{await S.db.doc("rh_nomina/"+per.id).update({estado:"borrador"});registrarActividad("Reabrió nómina",per.nombre);toast("Quincena reabierta.")}catch(e){toast("No se pudo reabrir.")}})},"Reabrir"):null)
      :h("div",{class:"status info"},"Borrador: se recalcula sola cada vez que cambias un empleado o una novedad. Ciérrala cuando la vayas a pagar para que quede fija.");
    const cerrar=async()=>{try{await S.db.doc("rh_nomina/"+per.id).set({periodo:per.id,mes:per.mes,q:per.q,desde:per.desde,hasta:per.hasta,estado:"cerrada",filas:f,totales:{devengado:T("devengado"),deducciones:T("deducciones"),neto:T("neto")},parametros:p,cerradaPor:S.uid,cerradaEn:nowISO()});
      registrarActividad("Cerró nómina",`${per.nombre} · ${f.length} empleados · neto ${fmt(T("neto"))}`);toast("Quincena cerrada.")}catch(e){toast("No se pudo cerrar: "+(e?.message||""))}};
    const aExcel=()=>excel(`nomina-${per.id}.xlsx`,[["Nómina",[["Empleado","Documento","Cargo","Salario mensual","Días liquidados","Días trabajados","Básico","Vacaciones y licencias pagadas","Incapacidades","Auxilio de transporte","Horas extra y recargos","Comisiones","Pagos no salariales","Total devengado","Salud 4 %","Pensión 4 %","Fondo de solidaridad","Otros descuentos","Total deducciones","Neto a pagar","Novedades"],
      ...f.map(L=>[L.nombre,L.documento,L.cargo,L.salario,L.dias,L.diasTrab,L.basico,L.ausPag,L.incap,L.aux,L.extras,L.comis,L.noSal,L.devengado,L.salud,L.pension,L.fsp,L.desc,L.deducciones,L.neto,(L.detalle||[]).join("; ")]),
      ["TOTAL","","","","","",T("basico"),T("ausPag"),T("incap"),T("aux"),T("extras"),T("comis"),T("noSal"),T("devengado"),T("salud"),T("pension"),T("fsp"),T("desc"),T("deducciones"),T("neto"),""]]]]);
    const aConciliacion=()=>{const I=h("input",{id:"nc-fecha",type:"date",value:per.hasta});let c=null;const b=h("button",{class:"pri",onclick:async()=>{if(!I.value)return;b.disabled=true;b.textContent="Enviando…";
        try{for(const L of f){if(L.neto>0)await crearRegistro("pago",{fecha:I.value,valor:L.neto,tercero:L.nombre,nit:L.documento,concepto:"Nómina",detalle:per.nombre,movId:null,sinComprobante:true,soporteId:null,origen:"rrhh:"+per.id})}
          await S.db.doc("rh_nomina/"+per.id).update({enviada:{fecha:I.value,por:S.uid,en:nowISO()}});c();toast("Pagos de nómina registrados en Conciliación.")}catch(e){toast("No se pudo completar el envío. Revisa Conciliación → Pagos antes de reintentar.");b.disabled=false;b.textContent="Reintentar"}}},"Registrar pagos");
      c=abrirModal("Enviar la nómina a Conciliación",h("div",{class:"stack"},h("p",{class:"muted",style:"margin:0"},`Se registra un pago por cada empleado (${f.filter(L=>L.neto>0).length}) con concepto «Nómina» y su neto a pagar. Cuando cargues el banco, cada transferencia se cruza sola con su empleado.`),campo("Fecha en que se paga",I,true),h("div",{class:"row"},b)),{ancho:520})};
    this.box.replaceChildren(estado,
      h("div",{class:"kpis"},kpi("Empleados",String(f.length),`${fmtF(per.desde)} a ${fmtF(per.hasta)}`),kpi("Total devengado",fmt(T("devengado"))),kpi("Deducciones",fmt(T("deducciones")),"salud, pensión y descuentos"),kpi("Neto a pagar",fmt(T("neto")),"lo que sale del banco")),
      h("div",{class:"card"},h("div",{class:"tbl"},h("table",null,h("thead",null,h("tr",null,h("th",null,"Empleado"),h("th",{class:"v"},"Días"),h("th",{class:"v"},"Básico"),h("th",{class:"v"},"Aux. transporte"),h("th",{class:"v"},"Extras"),h("th",{class:"v"},"Otros pagos"),h("th",{class:"v"},"Devengado"),h("th",{class:"v"},"Salud y pensión"),h("th",{class:"v"},"Otros desc."),h("th",{class:"v"},"Neto a pagar"))),
        h("tbody",null,f.map(L=>h("tr",{class:"click",title:"Ver el desprendible",onclick:()=>{const d=h("div",{class:"papel"});d.innerHTML=desprendible(L,per);abrirModal("Desprendible de "+L.nombre,h("div",{class:"stack"},d,h("div",{class:"row"},h("button",{class:"pri",onclick:()=>imprimir(`Nomina ${per.q}_${per.mes} ${L.nombre}`,[desprendible(L,per)])},"Imprimir o guardar en PDF"),h("button",{onclick:()=>descargarWord(`Nomina ${per.q}_${per.mes} ${L.nombre}`.replace(/[\\/:*?"<>|]/g,""),[desprendible(L,per)])},"Descargar en Word"))),{ancho:820})}},
          h("td",null,h("strong",null,L.nombre),L.detalle?.length?h("div",{class:"small muted"},L.detalle.join(" · ")):null),h("td",{class:"v"},String(L.diasTrab)+(L.diasTrab!==L.dias?` de ${L.dias}`:"")),h("td",{class:"v"},fmt(L.basico)),h("td",{class:"v"},fmt(L.aux)),h("td",{class:"v"},fmt(L.extras)),h("td",{class:"v"},fmt(L.ausPag+L.incap+L.comis+L.noSal)),
          h("td",{class:"v"},fmt(L.devengado)),h("td",{class:"v neg"},fmt(L.salud+L.pension+L.fsp)),h("td",{class:"v neg"},fmt(L.desc)),h("td",{class:"v"},h("strong",null,fmt(L.neto)))))),
        h("tfoot",null,h("tr",null,h("th",null,"Total"),h("th",null,""),h("th",{class:"v"},fmt(T("basico"))),h("th",{class:"v"},fmt(T("aux"))),h("th",{class:"v"},fmt(T("extras"))),h("th",{class:"v"},fmt(T("ausPag")+T("incap")+T("comis")+T("noSal"))),h("th",{class:"v"},fmt(T("devengado"))),h("th",{class:"v"},fmt(T("salud")+T("pension")+T("fsp"))),h("th",{class:"v"},fmt(T("desc"))),h("th",{class:"v"},fmt(T("neto")))))))),
      h("div",{class:"row"},h("button",{class:"pri",onclick:aExcel},"Descargar nómina en Excel"),h("button",{onclick:()=>imprimir("Desprendibles de nómina "+per.id,f.map(L=>desprendible(L,per)))},"Desprendibles de todos (imprimir o PDF)"),
        N.cerrada?(esAdmin()&&!N.doc.enviada?h("button",{onclick:aConciliacion},"Enviar pagos a Conciliación"):null):h("button",{onclick:ev=>confirmarEnPagina(ev.currentTarget,"Confirmar cierre",cerrar)},"Cerrar quincena")),
      h("div",{class:"small muted"},`Valores usados: salario mínimo ${fmt(p.smmlv)}, auxilio de transporte ${fmt(p.auxTransporte)} (para salarios hasta 2 mínimos), salud 4 % y pensión 4 % a cargo del empleado. Se cambian en Ajustes. La retención en la fuente y la liquidación de prestaciones no se calculan aquí.`));
  }
};

/* ---------- seguridad social (Asopagos) ---------- */
const TIPO_PLANILLA=["Tipo documento","Número documento","Primer apellido","Segundo apellido","Primer nombre","Segundo nombre","Cargo","Salario básico","Días cotizados","IBC","ING","Fecha ingreso","RET","Fecha retiro","VSP","Fecha cambio salario","SLN","Días SLN","Inicio SLN","Fin SLN","IGE","Días IGE","Inicio IGE","Fin IGE","LMA","Días LMA","Inicio LMA","Fin LMA","VAC-LR","Días VAC-LR","Inicio VAC-LR","Fin VAC-LR","IRL","Días IRL","Inicio IRL","Fin IRL","EPS","Fondo de pensiones","ARL","Clase de riesgo","Caja de compensación"];
function filaPlanilla(x){const e=x.e,n=x.nov,blo=c=>n[c]?["X",n[c].dias,fmtF(n[c].ini),fmtF(n[c].fin)]:["","","",""];
  return [e.tipoDoc||"CC",e.documento,e.primerApellido,e.segundoApellido||"",e.primerNombre,e.segundoNombre||"",e.cargo||"",x.sal,x.diasCot,x.ibc,x.ing?"X":"",fmtF(x.ing),x.ret?"X":"",fmtF(x.ret),n.VSP?"X":"",n.VSP?fmtF(n.VSP.ini):"",...blo("SLN"),...blo("IGE"),...blo("LMA"),...blo("VAC"),...blo("IRL"),e.eps||"",e.afp||"",e.arl||"",e.riesgo||"1",e.caja||""]}
const COD_NOV={ING:"Ingreso",RET:"Retiro",VSP:"Cambio de salario",SLN:"Suspensión o licencia no remunerada",IGE:"Incapacidad por enfermedad",LMA:"Licencia de maternidad o paternidad",VAC:"Vacaciones o licencia remunerada",IRL:"Incapacidad por accidente de trabajo"};
VIEWS.rh_seguridad={
  mount(el){iniciar();
    this.previa=h("div");this.area=h("div",{class:"stack"});this.cargadas=h("div");
    const fi=h("input",{type:"file",id:"ss-file",accept:".xlsx,.xls,.csv,.txt",onchange:()=>this.leer(fi.files[0])});
    el.append(h("div",null,h("h2",{id:"rh-st"},"Seguridad social"),h("div",{class:"small muted"},"Dos pasos: primero descargas el Excel del mes con los empleados y sus novedades para liquidar la planilla en Asopagos; después de pagarla, subes aquí la planilla para dejar el registro y revisar que no falte nadie.")),
      h("div",{class:"card stack"},h("div",{class:"row between"},h("h3",{style:"margin:0"},"1. Planilla del mes para Asopagos"),h("button",{class:"pri",onclick:()=>this.descargar()},"Descargar Excel para Asopagos")),this.previa),
      h("div",{class:"card stack"},h("h3",{style:"margin:0"},"2. Cargar la planilla pagada"),h("p",{class:"small muted",style:"margin:0"},"Sube el Excel que descargas de Asopagos con la liquidación (el detalle por cotizante). La app compara los cotizantes con tus empleados y guarda el valor pagado."),
        h("div",{class:"drop"},h("label",{class:"f",for:"ss-file"},"Archivo de Asopagos (Excel o CSV)",fi)),this.area),
      h("div",{class:"card"},h("h3",{id:"rh-sc"},"Planillas cargadas"),this.cargadas));
  },
  descargar(){const l=pilaMes(S.mes);if(!l.length){toast("No hay empleados activos en este mes.");return}
    excel(`planilla-asopagos-${S.mes}.xlsx`,[["Planilla",[TIPO_PLANILLA,...l.map(filaPlanilla)]],["Códigos",[["Código","Significado"],...Object.entries(COD_NOV)]]]);registrarActividad("Descargó planilla para Asopagos",`${mesNombre(S.mes)} · ${l.length} cotizantes`)},
  async leer(file){
    this.area.replaceChildren();if(!file)return;this.area.append(h("div",{class:"status info"},"Leyendo archivo…"));
    try{let rows;const ext=(file.name.split(".").pop()||"").toLowerCase();
      if(ext==="xlsx"||ext==="xls"){await cargarScript(XLSX_URL);const wb=XLSX.read(new Uint8Array(await file.arrayBuffer()),{type:"array"});rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,raw:true,defval:""})}
      else{const buf=await file.arrayBuffer();let txt;try{txt=new TextDecoder("utf-8",{fatal:true}).decode(buf)}catch(_){txt=new TextDecoder("windows-1252").decode(buf)}rows=parseCSV(txt)}
      rows=rows.filter(r=>r.some(c=>String(c).trim()!==""));if(!rows.length)throw new Error("vacío");
      const kw=/identificaci|documento|c[eé]dula|nombre|apellido|ibc|salud|pensi|riesgo|arl|caja|ccf|total|d[ií]as/i;
      let hi=-1;for(let i=0;i<Math.min(rows.length,40);i++){if(rows[i].filter(c=>kw.test(String(c))).length>=3){hi=i;break}}
      const ncol=Math.max(...rows.map(r=>r.length)),headers=Array.from({length:ncol},(_,i)=>hi>=0&&String(rows[hi][i]??"").trim()?String(rows[hi][i]).trim():`Columna ${i+1}`),data=rows.slice(hi+1),m=data.slice(0,30);
      const num=i=>m.filter(r=>String(r[i]??"").trim()!==""&&parseMonto(r[i])!=null).length>m.length/2;
      const find=(re,n)=>headers.findIndex((x,i)=>re.test(x)&&(n==null||num(i)===n));
      let doc=find(/(n[uú]mero|no\.?|nro\.?)\s*(de\s*)?(ident|doc)/i);if(doc<0)doc=find(/identificaci|documento|c[eé]dula/i,true);if(doc<0)doc=find(/identificaci|documento|c[eé]dula/i);
      this.st={archivo:file.name,headers,data,map:{doc,nombre:find(/nombre|apellido|cotizante/i,false),total:find(/total/i,true)},
        aportes:headers.map((x,i)=>/salud|eps|pensi|afp|riesgo|arl|caja|ccf|sena|icbf|solidaridad/i.test(x)&&!/ibc|d[ií]as|tarifa|c[oó]digo|nombre|administradora/i.test(x)&&num(i)?i:-1).filter(i=>i>=0)};
      this.pintarCarga();
    }catch(e){this.area.replaceChildren(h("div",{class:"status bad"},"No se pudo leer el archivo. Verifica que sea el Excel (.xlsx) o CSV que descargaste de Asopagos."))}
  },
  filasCarga(){const st=this.st,m=st.map;return st.data.map(r=>({doc:soloNum(r[m.doc]),nombre:m.nombre>=0?String(r[m.nombre]??"").trim():"",total:Math.round(m.total>=0?Math.abs(parseMonto(r[m.total])||0):st.aportes.reduce((s,i)=>s+Math.abs(parseMonto(r[i])||0),0))})).filter(x=>x.doc.length>=5)},
  pintarCarga(){
    const st=this.st,I={};
    const sel=(k,l,opc)=>campo(l,h("select",{id:"ss-map-"+k,onchange:e=>{st.map[k]=+e.target.value;this.pintarCarga()}},opc?h("option",{value:"-1"},opc):null,st.headers.map((x,i)=>h("option",{value:String(i),selected:st.map[k]===i?true:null},x))));
    if(st.map.doc<0)return this.area.replaceChildren(h("div",{class:"status warn"},"No se encontró la columna del número de documento. Elígela:"),sel("doc","Número de documento"));
    const filas=this.filasCarga(),total=suma(filas,"total"),ini=S.mes+"-01",fin=S.mes+"-31";
    const act=RH.emp.filter(e=>activoEn(e,ini,fin)),enPl=new Set(filas.map(f=>f.doc)),misDocs=new Set(RH.emp.map(e=>soloNum(e.documento)));
    const faltan=act.filter(e=>!enPl.has(soloNum(e.documento))),sobran=filas.filter(f=>!misDocs.has(f.doc));
    const btn=h("button",{class:"pri",disabled:filas.length?null:true,onclick:async()=>{btn.disabled=true;
      const tot=Math.abs(parseMonto(I.total.value)||0);
      try{await S.db.doc(`rh_planillas/${S.mes}_${rid()}`).set({mes:S.mes,numero:I.numero.value.trim(),fechaPago:I.fecha.value,total:tot,archivo:st.archivo,cotizantes:filas.length,filas:filas.slice(0,3000),faltan:faltan.map(nombreEmp),createdBy:S.uid,createdAt:nowISO()});
        registrarActividad("Cargó planilla de seguridad social",`${mesNombre(S.mes)} · ${filas.length} cotizantes · ${fmt(tot)}`);toast("Planilla guardada.");this.area.replaceChildren(h("div",{class:"status ok"},"Planilla guardada. Aparece abajo en «Planillas cargadas»."));$("#ss-file").value=""}
      catch(e){toast("No se pudo guardar la planilla.");btn.disabled=false}}},"Guardar planilla de "+mesNombre(S.mes).toLowerCase());
    this.area.replaceChildren(h("div",{class:"status ok"},`Archivo «${st.archivo}»: ${filas.length} cotizantes. Confirma las columnas.`),
      h("div",{class:"cols3"},sel("doc","Número de documento"),sel("nombre","Nombre",`— ninguna —`),sel("total","Total pagado por cotizante","— sumar las columnas de aportes —")),
      h("div",{class:"kpis"},kpi("Cotizantes en la planilla",String(filas.length)),kpi("Total de la planilla",fmt(total)),kpi("Empleados que faltan",String(faltan.length),faltan.length?"activos este mes y no están en la planilla":"todos los activos están",faltan.length>0)),
      faltan.length?h("div",{class:"status bad"},"No aparecen en la planilla: "+faltan.map(nombreEmp).join(", ")+"."):null,
      sobran.length?h("div",{class:"status warn"},`En la planilla hay ${sobran.length} ${sobran.length===1?"persona que no está":"personas que no están"} en Empleados: `+sobran.slice(0,12).map(f=>f.nombre||f.doc).join(", ")+(sobran.length>12?"…":"")+"."):null,
      h("div",{class:"cols3"},campo("Número de la planilla",I.numero=h("input",{id:"ss-numero",autocomplete:"off"})),campo("Fecha de pago",I.fecha=h("input",{id:"ss-fecha",type:"date",value:hoy()})),campo("Valor total pagado",I.total=h("input",{id:"ss-total",inputmode:"decimal",value:String(total)}),false,"Corrígelo si incluye intereses de mora.")),
      h("div",{class:"row"},btn));
  },
  refresh(){
    $("#rh-st").textContent="Seguridad social · "+mesNombre(S.mes);$("#rh-sc").textContent="Planillas cargadas de "+mesNombre(S.mes).toLowerCase();
    if(cargandoRH()){this.previa.replaceChildren(tablaVacia("Cargando…"));return}
    const l=pilaMes(S.mes);
    this.previa.replaceChildren(l.length?h("div",{class:"stack"},h("div",{class:"small muted"},`${l.length} cotizantes · base de cotización total ${fmt(suma(l,"ibc"))}. Las novedades salen de la pestaña Novedades y de los ingresos y retiros de Empleados.`),
      h("div",{class:"tbl",style:"max-height:380px;overflow:auto"},h("table",null,h("thead",null,h("tr",null,h("th",null,"Empleado"),h("th",{class:"v"},"Salario"),h("th",{class:"v"},"Días"),h("th",{class:"v"},"Base (IBC)"),h("th",null,"Novedades del mes"))),
        h("tbody",null,l.map(x=>{const ns=[x.ing&&["ING","el "+fmtF(x.ing)],x.ret&&["RET","el "+fmtF(x.ret)],...Object.entries(x.nov).map(([c,v])=>[c,v.dias?`${v.dias} días`:"desde "+fmtF(v.ini)])].filter(Boolean);
          return h("tr",null,h("td",null,h("strong",null,nombreEmp(x.e)),h("div",{class:"small muted"},`${x.e.tipoDoc||"CC"} ${x.e.documento}`)),h("td",{class:"v"},fmt(x.sal)),h("td",{class:"v"},String(x.diasCot)),h("td",{class:"v"},fmt(x.ibc)),
            h("td",null,ns.length?h("div",{class:"row",style:"gap:6px"},ns.map(([c,t])=>h("span",{class:"chip "+(c==="ING"?"ok":c==="RET"?"bad":"info"),title:COD_NOV[c]},`${COD_NOV[c]}: ${t}`))):h("span",{class:"muted small"},"Sin novedades")))})))))
      :tablaVacia(RH.emp.length?"No hay empleados activos en este mes.":"Aún no hay empleados. Créalos en la pestaña Empleados."));
    const ps=RH.planillas.filter(p=>p.mes===S.mes).sort((a,b)=>String(b.createdAt).localeCompare(String(a.createdAt)));
    this.cargadas.replaceChildren(ps.length?h("div",{class:"tbl"},h("table",null,h("thead",null,h("tr",null,h("th",null,"Planilla"),h("th",{class:"v"},"Cotizantes"),h("th",{class:"v"},"Valor pagado"),h("th",null,"Fecha de pago"),h("th",null,"Cargó"),h("th",null,""))),
      h("tbody",null,ps.map(p=>h("tr",null,h("td",null,h("strong",null,p.numero||"Sin número"),h("div",{class:"small muted"},p.archivo),p.faltan?.length?h("div",{class:"small neg"},"Faltaron: "+p.faltan.join(", ")):null),h("td",{class:"v"},String(p.cotizantes)),h("td",{class:"v"},fmt(p.total)),h("td",null,fmtF(p.fechaPago)),h("td",{class:"small"},nombre(p.createdBy),h("div",{class:"muted"},fmtFH(p.createdAt))),
        h("td",null,h("div",{class:"row",style:"gap:6px;flex-wrap:nowrap"},
          esAdmin()?(p.enviada?h("span",{class:"chip ok"},"En Conciliación"):h("button",{class:"sm pri",title:"Registra el pago con concepto Seguridad social para que se cruce con el banco",onclick:async ev=>{const b=ev.currentTarget;if(!p.total||!p.fechaPago){toast("La planilla no tiene valor o fecha de pago.");return}b.disabled=true;
            try{await crearRegistro("pago",{fecha:p.fechaPago,valor:p.total,tercero:"Asopagos (planilla de seguridad social)",nit:"",concepto:"Seguridad social",detalle:`Planilla ${p.numero||""} · ${mesNombre(p.mes)}`.trim(),movId:null,sinComprobante:true,soporteId:null,origen:"rrhh:planilla:"+p.id});
              await S.db.doc("rh_planillas/"+p.id).update({enviada:{por:S.uid,en:nowISO()}});toast("Pago registrado en Conciliación.")}catch(e){toast("No se pudo registrar el pago.");b.disabled=false}}},"Enviar a Conciliación")):null,
          h("button",{class:"sm danger",onclick:ev=>confirmarEnPagina(ev.currentTarget,"Confirmar",async()=>{try{await S.db.doc("rh_planillas/"+p.id).delete();registrarActividad("Eliminó planilla de seguridad social",`${p.numero||p.archivo}`);toast("Planilla eliminada.")}catch(e){toast("No se pudo eliminar.")}})},"Eliminar"))))))))
      :tablaVacia("Aún no se ha cargado la planilla pagada de este mes."));
  }
};

/* ---------- contratos y documentos ---------- */
VIEWS.rh_documentos={
  mount(el){iniciar();
    this.hist=h("div");
    el.append(h("div",null,h("h2",null,"Contratos y documentos"),h("div",{class:"small muted"},"Elige el documento, elige el empleado y la app lo llena con sus datos. Queda listo para imprimir, guardar en PDF o abrir en Word.")),
      h("div",{class:"card doc-card",style:"border-color:var(--accent)"},h("h3",null,"¿Ingresa alguien nuevo?"),h("p",{class:"muted"},"Crea al empleado y al guardar se generan su contrato y su autorización de datos, listos para firmar."),h("button",{class:"pri",onclick:()=>formEmpleado(null)},"Nuevo empleado")),
      h("div",{class:"doc-grid"},Object.entries(DOCS).map(([k,d])=>h("div",{class:"card doc-card"},h("h3",null,d.l),h("p",{class:"small muted"},d.d),h("button",{id:"doc-"+k,onclick:()=>generador(k)},"Generar")))),
      h("div",{class:"card"},h("h3",null,"Documentos generados"),this.hist));
  },
  refresh(){
    const l=RH.docs.slice().sort((a,b)=>String(b.createdAt).localeCompare(String(a.createdAt))).slice(0,200);
    this.hist.replaceChildren(l.length?h("div",{class:"tbl",style:"max-height:460px;overflow:auto"},h("table",null,h("thead",null,h("tr",null,h("th",null,"Fecha"),h("th",null,"Documento"),h("th",null,"Empleado"),h("th",null,"Generó"),h("th",null,""))),
      h("tbody",null,l.map(d=>h("tr",null,h("td",{class:"small",style:"white-space:nowrap"},fmtFH(d.createdAt)),h("td",null,d.titulo),h("td",null,d.empleado||"—"),h("td",null,nombre(d.createdBy)),
        h("td",null,DOCS[d.tipo]&&empPor(d.empleadoId)?h("button",{class:"sm",onclick:()=>generador(d.tipo,{empleadoId:d.empleadoId,campos:d.campos})},"Volver a generar"):null))))))
      :tablaVacia("Aún no se ha generado ningún documento."));
  }
};

/* ---------- ajustes de recursos humanos ---------- */
VIEWS.rh_ajustes={
  mount(el){iniciar();this.lleno=false;
    const E={},V={};this.E=E;this.V=V;
    const fe=(k,l,ph)=>campo(l,E[k]=h("input",{id:"aj-"+k,autocomplete:"off",placeholder:ph||null}));
    const memb=E.membrete=h("input",{type:"checkbox",id:"aj-membrete",checked:true});
    const bE=h("button",{class:"pri",type:"submit"},"Guardar datos de la empresa");
    const formE=h("form",{class:"stack",onsubmit:async ev=>{ev.preventDefault();bE.disabled=true;
      const d={};for(const k of Object.keys(E))d[k]=k==="membrete"?memb.checked:E[k].value.trim();
      try{await S.db.doc("rh_config/empresa").set({...d,actualizadoPor:S.uid,actualizadoEn:nowISO()});registrarActividad("Actualizó datos de la empresa para documentos","");toast("Datos de la empresa guardados.")}catch(e){toast("No se pudo guardar.")}bE.disabled=false}},
      h("div",{class:"cols3"},fe("razonSocial","Razón social","Nombre legal de la empresa"),fe("nit","NIT","900.123.456-7"),fe("direccion","Dirección"),fe("ciudad","Ciudad","Ej. Medellín"),fe("telefono","Teléfono"),
        fe("representante","Representante legal"),fe("documentoRep","Cédula del representante legal"),fe("firmante","Quien firma las cartas y certificados","Si se deja vacío, el representante legal"),fe("cargoFirmante","Cargo de quien firma","Representante legal")),
      h("label",{class:"row",for:"aj-membrete",style:"gap:8px"},memb,"Poner el nombre y los datos de la empresa como encabezado de cada documento"),h("div",{class:"row"},bE));
    const fv=(k,l,ay)=>campo(l,V[k]=h("input",{id:"aj-"+k,inputmode:"decimal"}),false,ay);
    const bV=h("button",{class:"pri",type:"submit"},"Guardar valores");
    const formV=h("form",{class:"stack",onsubmit:async ev=>{ev.preventDefault();const d={};
      for(const k of Object.keys(V)){const n=Math.abs(parseMonto(V[k].value)||0);if(!n){toast("Todos los valores deben ser mayores que cero.");return}d[k]=n}
      bV.disabled=true;try{await S.db.doc("rh_config/parametros").set({...d,actualizadoPor:S.uid,actualizadoEn:nowISO()});registrarActividad("Actualizó valores legales de nómina",`Mínimo ${fmt(d.smmlv)}`);toast("Valores guardados. La nómina en borrador ya los usa.")}catch(e){toast("No se pudo guardar.")}bV.disabled=false}},
      h("div",{class:"cols3"},fv("smmlv","Salario mínimo mensual"),fv("auxTransporte","Auxilio de transporte mensual"),fv("horasMes","Horas de trabajo al mes","Para el valor de la hora extra. Con jornada de 42 horas semanales son 210."),fv("recargoDominical","Recargo dominical y festivo (%)","Sube por etapas con la reforma laboral: revísalo cada julio.")),
      h("div",{class:"row"},bV));
    /* plantillas */
    const ta=h("textarea",{id:"pl-texto",class:"plantilla",spellcheck:"true"}),prev=h("div",{class:"papel",hidden:true}),estado=h("span",{class:"small muted"});
    const sel=h("select",{id:"pl-sel",style:"max-width:340px",onchange:()=>cargar()},Object.entries(PLANT).map(([k,p])=>h("option",{value:k},p.l)));
    const cargar=()=>{ta.value=textoPlantilla(sel.value);prev.hidden=true;estado.textContent=RH.plantillas[sel.value]?.cuerpo?`Texto propio, guardado por ${nombre(RH.plantillas[sel.value].actualizadoPor)} el ${fmtFH(RH.plantillas[sel.value].actualizadoEn)}`:"Texto original de la app"};
    this.cargarPlantilla=cargar;this.ta=ta;
    const chips=h("div",{class:"campos-chips"},CAMPOS.map(([k,l])=>h("button",{type:"button",title:l,onclick:()=>{ta.focus();ta.setRangeText(`{{${k}}}`,ta.selectionStart,ta.selectionEnd,"end")}},l)));
    el.append(h("div",null,h("h2",null,"Ajustes de Recursos Humanos"),h("div",{class:"small muted"},"Se configuran una sola vez y se usan en la nómina y en todos los documentos.")),
      h("div",{class:"card"},h("h3",null,"Datos de la empresa para los documentos"),formE),
      h("div",{class:"card"},h("h3",null,"Valores legales del año"),h("p",{class:"small muted",style:"margin:0 0 10px"},"Vienen con los valores de 2026. Verifícalos con el decreto del Gobierno y actualízalos cada enero."),formV),
      h("div",{class:"card stack"},h("h3",{style:"margin:0"},"Plantillas de contratos y cartas"),
        h("p",{class:"small muted",style:"margin:0"},"Aquí pegas tu propio borrador. Donde va un dato del empleado o de la empresa, toca el botón del dato y la app pone el campo entre llaves, por ejemplo {{nombre}}. Una línea que empieza con # es el título, **texto** queda en negrilla y [[FIRMAS]] pone las firmas de empleador y trabajador ([[FIRMA_EMPRESA]], [[FIRMA_TRABAJADOR]] y [[RECIBIDO]] ponen una sola). Deja una línea en blanco entre párrafos."),
        h("div",{class:"row"},sel,estado),chips,ta,
        h("div",{class:"row"},h("button",{class:"pri",onclick:async ev=>{const b=ev.currentTarget;if(ta.value.trim().length<40){toast("El texto está vacío o es muy corto.");return}b.disabled=true;
            try{await S.db.doc("rh_plantillas/"+sel.value).set({cuerpo:ta.value,actualizadoPor:S.uid,actualizadoEn:nowISO()});registrarActividad("Cambió plantilla de documento",PLANT[sel.value].l);toast("Plantilla guardada.")}catch(e){toast("No se pudo guardar.")}b.disabled=false}},"Guardar plantilla"),
          h("button",{onclick:()=>{const e=RH.emp.slice().sort(porNombre)[0]||null;const c=cuerpoDoc(ta.value,valores(e,{fecha_desde:hoy(),fecha_hasta:addDias(hoy(),2),motivo:"(aquí va la causa)",dirigido_a:"A quien interese",fecha_retiro:hoy()}));prev.hidden=false;prev.innerHTML=`<div class="hoja">${c.html}</div>`;prev.scrollIntoView({block:"nearest",behavior:"smooth"})}},"Vista previa"),
          h("button",{class:"danger",onclick:ev=>confirmarEnPagina(ev.currentTarget,"Confirmar",async()=>{try{if(RH.plantillas[sel.value])await S.db.doc("rh_plantillas/"+sel.value).delete();delete RH.plantillas[sel.value];cargar();toast("Se restauró el texto original.")}catch(e){toast("No se pudo restaurar.")}})},"Volver al texto original")),
        prev,h("div",{class:"status info"},"Los textos originales son modelos generales basados en el Código Sustantivo del Trabajo. Antes de usarlos con tus empleados, hazlos revisar por tu abogado laboral y ajústalos a tu empresa.")));
    cargar();
  },
  refresh(){
    if(this.lleno||!RH.ok.cfg)return;this.lleno=true;const p=P();
    for(const k of Object.keys(this.E)){if(k==="membrete")this.E[k].checked=RH.empresa.membrete!==false;else this.E[k].value=RH.empresa[k]||(k==="razonSocial"?nombreEmpresa():"")}
    for(const k of Object.keys(this.V))this.V[k].value=String(p[k]);
    if(document.activeElement!==this.ta)this.cargarPlantilla();
  }
};

/* ---------- alertas de recursos humanos (aparecen arriba, como las de conciliación) ---------- */
window.rhAlertas=()=>{
  const it=[];if(!RH.ok.emp||!puedeRH())return it;const hf=hoy();
  for(const e of RH.emp){
    if(e.tipoContrato!=="fijo"||!e.fechaFin||e.fechaRetiro)continue;const d=diasEntre(hf,e.fechaFin);if(d<0||d>45)continue;
    if(RH.docs.some(x=>x.empleadoId===e.id&&x.tipo==="finalizacion"&&x.campos?.fecha_fin===e.fechaFin))continue;
    it.push({c:d>=30?"warn":"",t:`El contrato a término fijo de ${nombreEmp(e)} vence el ${fmtF(e.fechaFin)}`,
      s:d>=30?`Faltan ${d} días. Si no se va a renovar, entrega el preaviso a más tardar el ${fmtF(addDias(e.fechaFin,-30))}. Toca para generarlo.`:`Faltan ${d} días y ya pasó el plazo de 30 días para el preaviso: si no hay acuerdo con el empleado, el contrato se prorroga. Actualiza la fecha en su ficha si se renueva.`,
      fn:()=>generador("finalizacion",{empleadoId:e.id})})}
  if(RH.ok.cfg&&RH.emp.length&&!RH.empresa.razonSocial&&S.tab!=="rh_ajustes")it.push({c:"warn",t:"Faltan los datos de la empresa para los documentos",s:"NIT, dirección, ciudad y representante legal salen en contratos y certificados. Toca para completarlos.",tab:"rh_ajustes"});
  return it;
};
})();
