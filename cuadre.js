/* =====================================================================
   MÓDULO: CUADRE DE INVENTARIO (solo productos con stock en el sistema)
   - Solo lectura: no escribe ni borra nada en Firebase.
   - Usa la cantidad final (si hay reconteo, manda el reconteo).
   - Excel: un producto con varios lotes / vencimientos sale en varias
     filas, una debajo de otra.
   Se carga DESPUÉS del script principal de index.html:
   <script src="cuadre.js"></script>
   ===================================================================== */
(function(){
'use strict';

var cFiltro = 'con';      // tarjeta seleccionada
var cTimer  = null;

// ---------- estilos ----------
var css = document.createElement('style');
css.textContent =
'#mCuadre .mbox{max-width:1300px;width:100%;height:94vh;max-height:94vh;display:flex;flex-direction:column;padding:16px}'+
'#cTop{display:flex;justify-content:space-between;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:10px}'+
'#cCards{display:grid;grid-template-columns:repeat(auto-fit,minmax(120px,1fr));gap:8px;margin-bottom:10px}'+
'#cCards div{border:1px solid var(--line);border-radius:10px;padding:8px 10px;cursor:pointer;background:#fff}'+
'#cCards div.on{outline:2px solid var(--pri)}'+
'#cCards b{display:block;font-size:20px}#cCards span{font-size:12px;color:var(--mut)}'+
'#cFil{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:8px}'+
'#cFil input,#cFil select{padding:8px;border:1px solid var(--line);border-radius:8px;width:auto}'+
'#cWrap{flex:1;overflow:auto;border:1px solid var(--line);border-radius:10px}'+
'#cTabla{width:100%;border-collapse:collapse;font-size:13px}'+
'#cTabla th{position:sticky;top:0;background:var(--pri);color:#fff;padding:7px 6px;text-align:left;white-space:nowrap;z-index:1}'+
'#cTabla td{padding:6px;border-bottom:1px solid var(--line);vertical-align:top}'+
'#cTabla td.n{text-align:right;white-space:nowrap}'+
'#cTabla tr.sub td{border-top:none;background:#fafcfc}'+
'#cTabla tr.sub td.vac{border-bottom-color:transparent}'+
'.cF{color:var(--red);font-weight:600}.cS{color:#1565c0;font-weight:600}.cOk{color:var(--ok);font-weight:600}.cP{color:var(--warn);font-weight:600}';
document.head.appendChild(css);

// ---------- modal ----------
var m = document.createElement('div');
m.className = 'modal'; m.id = 'mCuadre';
m.innerHTML =
'<div class="mbox">'+
 '<div id="cTop">'+
  '<h3 style="margin:0">✅ Cuadre de inventario — solo productos con stock en el sistema</h3>'+
  '<div class="row" style="gap:8px">'+
   '<button class="btn btn-p" onclick="cuadreExcel()">⬇ Descargar Excel</button>'+
   '<button class="btn" onclick="cuadreCerrar()">Cerrar</button>'+
  '</div>'+
 '</div>'+
 '<div id="cCards"></div>'+
 '<div id="cFil">'+
  '<input id="cQ" placeholder="Buscar código o descripción..." oninput="cuadreRender()" style="flex:1;min-width:200px">'+
  '<select id="cLab" onchange="cuadreRender()"></select>'+
 '</div>'+
 '<div id="cWrap"><table id="cTabla"><thead><tr>'+
  '<th>Laboratorio</th><th>Código</th><th>Descripción</th><th>Stock sistema</th><th>Contado total</th><th>Diferencia</th><th>Estado</th>'+
  '<th>Lote</th><th>Vcto.</th><th>Cant. lote</th>'+
 '</tr></thead><tbody id="cBody"></tbody></table></div>'+
 '<div class="sub" id="cPie" style="color:var(--mut);margin-top:6px"></div>'+
'</div>';
document.body.appendChild(m);

// ---------- botón en el panel de administrador ----------
function ponerBoton(){
  if(document.getElementById('btnCuadre')) return;
  var ref = document.querySelector('[onclick*="abrirReconteo"]') || document.querySelector('[onclick*="abrirGeneral"]');
  if(!ref) return;
  var b = document.createElement('button');
  b.id = 'btnCuadre'; b.className = 'btn';
  b.style.cssText = 'width:100%;margin-bottom:14px;background:#2a9d8f;color:#fff';
  b.textContent = '✅ Cuadre: contados, sobrantes, faltantes, cuadrados y pendientes (solo con stock)';
  b.onclick = function(){ window.cuadreAbrir(); };
  ref.parentNode.insertBefore(b, ref.nextSibling);
}
setTimeout(ponerBoton, 0);
document.addEventListener('click', function(){ setTimeout(ponerBoton, 50); }, true);

// ---------- cálculo ----------
var NOMBRE = {pen:'PENDIENTE', falta:'FALTANTE', sobra:'SOBRANTE', ok:'CUADRADO'};
var CLASE  = {pen:'cP', falta:'cF', sobra:'cS', ok:'cOk'};

function base(){
  var lab = document.getElementById('cLab').value;
  return lista().filter(function(p){ return conStock(p) && (!lab || p.lab === lab); });
}
function pasaFiltro(p, f){
  var e = estadoDe(p);
  if(f === 'todos') return true;
  if(f === 'con') return e !== 'pen';
  return e === f;
}
function filtrados(f){
  var q = (document.getElementById('cQ').value || '').trim().toUpperCase();
  return base().filter(function(p){
    if(q && !((p.cod||'').toUpperCase().indexOf(q) >= 0 || (p.desc||'').toUpperCase().indexOf(q) >= 0)) return false;
    return pasaFiltro(p, f);
  }).sort(function(a,b){ return a.lab.localeCompare(b.lab) || a.desc.localeCompare(b.desc); });
}
function stockTxt(p){ return txtCant(+p.stkE||0, +p.stkF||0); }

// ---------- pantalla ----------
window.cuadreAbrir = function(){
  if(!esAdmin(firebase.auth().currentUser)){ openAdmin(); return; }
  var labs = Array.from(new Set(lista().filter(conStock).map(function(p){ return p.lab; }))).sort();
  var sel = document.getElementById('cLab'), prev = sel.value;
  sel.innerHTML = '<option value="">Todos los laboratorios</option>' + labs.map(function(l){ return '<option>'+esc(l)+'</option>'; }).join('');
  sel.value = prev;
  if(typeof cerrar === 'function') cerrar('mAdmin');
  if(typeof escucharReco === 'function') escucharReco();   // trae reconteos: manda la cantidad final
  m.classList.add('on');
  cuadreRender();
  clearInterval(cTimer);
  cTimer = setInterval(function(){
    if(!m.classList.contains('on')){ clearInterval(cTimer); return; }
    if(document.activeElement && document.activeElement.id === 'cQ') return;
    cuadreRender();
  }, 4000);
};
window.cuadreCerrar = function(){ m.classList.remove('on'); clearInterval(cTimer); };
window.cuadreCard = function(v){ cFiltro = v; cuadreRender(); };

window.cuadreRender = function(){
  if(!m.classList.contains('on')) return;
  var B = base();
  var n = {todos:B.length, con:0, sobra:0, falta:0, ok:0, pen:0};
  B.forEach(function(p){ var e = estadoDe(p); n[e]++; if(e !== 'pen') n.con++; });
  var card = function(k, t){ return '<div class="'+(cFiltro===k?'on':'')+'" onclick="cuadreCard(\''+k+'\')"><b>'+n[k]+'</b><span>'+t+'</span></div>'; };
  document.getElementById('cCards').innerHTML =
    card('todos','Con stock en sistema') + card('con','Contados') + card('sobra','Sobrantes') +
    card('falta','Faltantes') + card('ok','Cuadrados') + card('pen','Pendientes');

  var rows = filtrados(cFiltro), MAX = 1500, html = [];
  rows.slice(0, MAX).forEach(function(p){
    var e = estadoDe(p), d = difDe(p);
    var lotes = contadoFin(p) ? lotesFinal(p) : [];
    if(!lotes.length) lotes = [null];
    var dif = d === null ? '<span class="cP">—</span>' : '<span class="'+(d<0?'cF':d>0?'cS':'cOk')+'">'+(d>0?'+':'')+fmtDif(p,d)+'</span>';
    var tot = contadoFin(p) ? txtCant(totalFin(p), totalFFin(p)) : '<span class="cP">—</span>';
    lotes.forEach(function(f, i){
      var lt = f ? '<td>'+esc(f.lote||'s/l')+'</td><td>'+esc(fmtVenc(f.venc))+'</td><td class="n">'+txtCant(f.cant, f.fr)+(f.reco?' 🔁':'')+'</td>'
                 : '<td colspan="3" class="cP">Sin contar</td>';
      if(i === 0){
        html.push('<tr><td>'+esc(p.lab)+'</td><td>'+esc(p.cod)+'</td><td>'+esc(p.desc)+'</td>'+
          '<td class="n">'+stockTxt(p)+'</td><td class="n">'+tot+'</td><td class="n">'+dif+'</td>'+
          '<td><span class="'+CLASE[e]+'">'+NOMBRE[e]+'</span></td>'+lt+'</tr>');
      } else {
        html.push('<tr class="sub"><td class="vac"></td><td class="vac">'+esc(p.cod)+'</td><td class="vac" style="color:var(--mut)">'+esc(p.desc)+'</td>'+
          '<td class="vac"></td><td class="vac"></td><td class="vac"></td><td class="vac"></td>'+lt+'</tr>');
      }
    });
  });
  document.getElementById('cBody').innerHTML = html.join('') ||
    '<tr><td colspan="10" style="text-align:center;padding:20px;color:var(--mut)">Sin resultados</td></tr>';
  document.getElementById('cPie').textContent =
    'Mostrando ' + Math.min(rows.length, MAX) + ' de ' + rows.length + ' productos con stock' +
    (rows.length > MAX ? ' (usa los filtros para ver el resto)' : '') +
    '. Si un producto tiene varios lotes o vencimientos, salen uno debajo de otro. 🔁 = cantidad del reconteo. Se actualiza solo.';
};

// ---------- Excel ----------
function filasExcel(lista_){
  var out = [];
  lista_.forEach(function(p){
    var e = estadoDe(p), d = difDe(p), cf = contadoFin(p), r = (typeof recoDe === 'function') ? recoDe(p.k) : null;
    var lotes = cf ? lotesFinal(p) : [];
    if(!lotes.length) lotes = [null];
    lotes.forEach(function(f, i){
      var o = {Laboratorio: p.lab, Codigo: p.cod, Descripcion: p.desc, Tipo: p.tipo || '', Frac: p.frac || 1};
      // Stock, total contado y diferencia del producto solo en su primera fila (para que no se sumen dos veces)
      o.Stk_E_Sistema = i === 0 ? (+p.stkE||0) : '';
      if(FR) o.Stk_F_Sistema = i === 0 ? (+p.stkF||0) : '';
      o.Contado_Total = i === 0 && cf ? totalFin(p) : '';
      if(FR) o.Contado_Total_Fr = i === 0 && cf ? totalFFin(p) : '';
      o.Diferencia = i === 0 && d !== null ? (FR ? fmtDif(p, d) : d) : '';
      o.Estado = NOMBRE[e];
      o.Lote = f ? (f.lote || 's/l') : '';
      o.Vencimiento = f ? fmtVenc(f.venc) : '';
      o.Cant_Lote = f ? f.cant : '';
      if(FR) o.Cant_Lote_Fr = f ? f.fr : '';
      o.Lote_N = f ? (i+1) + ' de ' + lotes.length : '';
      o.Reconteo_Lote = f && f.reco ? 'SI' : '';
      o.Antes_Reconteo = f && f.reco ? (f.antes === null || f.antes === undefined ? 'NO CONTADO' : txtCant(f.antes, f.antesF)) : '';
      o.Responsable = i === 0 && p.c ? (p.c.por || '') : '';
      o.Fecha_Conteo = i === 0 && p.c ? fmtFecha(p.c.fecha) : '';
      o.Recontó = i === 0 && r ? (r.por || '') : '';
      o.Obs_Reconteo = i === 0 && r ? (r.obs || '') : '';
      o.Comentario = i === 0 && p.c ? (p.c.coment || '') : '';
      out.push(o);
    });
  });
  return out;
}
function hoja(wb, nombre, datos){
  var ws = XLSX.utils.json_to_sheet(datos.length ? datos : [{Laboratorio:'(sin productos)'}]);
  var cols = Object.keys(datos[0] || {Laboratorio:1});
  ws['!cols'] = cols.map(function(c){ return {wch: c === 'Descripcion' ? 42 : c === 'Laboratorio' ? 28 : Math.max(10, c.length + 2)}; });
  ws['!autofilter'] = {ref: XLSX.utils.encode_range({s:{r:0,c:0}, e:{r:Math.max(datos.length,1), c:cols.length-1}})};
  XLSX.utils.book_append_sheet(wb, ws, nombre);
}
window.cuadreExcel = function(){
  if(typeof XLSX === 'undefined'){ alert('No cargó el generador de Excel. Revisa tu internet.'); return; }
  var B = base(), lab = document.getElementById('cLab').value;
  var ord = function(a,b){ return a.lab.localeCompare(b.lab) || a.desc.localeCompare(b.desc); };
  var de = function(f){ return B.filter(function(p){ return pasaFiltro(p, f); }).sort(ord); };
  var n = {con:0, sobra:0, falta:0, ok:0, pen:0};
  B.forEach(function(p){ var e = estadoDe(p); n[e]++; if(e !== 'pen') n.con++; });

  var wb = XLSX.utils.book_new();
  var res = XLSX.utils.aoa_to_sheet([
    ['CUADRE DE INVENTARIO — productos con stock en el sistema'],
    ['Almacén', (typeof ALM !== 'undefined' && ALM && (ALM.nombre || ALM.id)) || ''],
    ['Laboratorio', lab || 'Todos'],
    ['Generado', new Date().toLocaleString('es-PE')],
    [],
    ['Concepto', 'Productos'],
    ['Con stock en sistema', B.length],
    ['Contados', n.con],
    ['Sobrantes', n.sobra],
    ['Faltantes', n.falta],
    ['Cuadrados', n.ok],
    ['Pendientes', n.pen],
    ['% avance', B.length ? Math.round(n.con / B.length * 10000) / 100 + '%' : '0%']
  ]);
  res['!cols'] = [{wch:28},{wch:22}];
  XLSX.utils.book_append_sheet(wb, res, 'Resumen');
  hoja(wb, 'Todos con stock', filasExcel(de('todos')));
  hoja(wb, 'Contados',   filasExcel(de('con')));
  hoja(wb, 'Sobrantes',  filasExcel(de('sobra')));
  hoja(wb, 'Faltantes',  filasExcel(de('falta')));
  hoja(wb, 'Cuadrados',  filasExcel(de('ok')));
  hoja(wb, 'Pendientes', filasExcel(de('pen')));
  var f = new Date(), dd = String(f.getDate()).padStart(2,'0') + '-' + String(f.getMonth()+1).padStart(2,'0') + '-' + f.getFullYear();
  XLSX.writeFile(wb, 'Cuadre_inventario_' + dd + '.xlsx');
};
})();
