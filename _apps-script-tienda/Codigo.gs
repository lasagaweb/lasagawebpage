/**
 * Tienda de La Saga: cobro con Mercado Pago, aviso a Factoría y stock.
 *
 * Va pegado en un proyecto de Google Apps Script ligado a la planilla «Tienda La Saga».
 * La carpeta empieza con _ para que GitHub Pages no la publique.
 *
 * Planilla:
 *   Hoja «Stock»:   id | nombre | desc | img | precio | S | M | L | XL | XXL
 *   Hoja «Pedidos»: la arma este script (una fila por pedido).
 *
 * Propiedades del script (Configuración del proyecto › Propiedades del script):
 *   MP_ACCESS_TOKEN   token de producción de la cuenta de Mercado Pago de La Saga (lo carga Nico)
 *   MAIL_FACTORIA     contacto@factoriamerch.com
 *   MAIL_BANDA        mail de La Saga que recibe copia
 *   URL_TIENDA        https://lasaga.co/tienda
 */

const TALLES = ['S', 'M', 'L', 'XL', 'XXL'];
const ENVIO_MONTEVIDEO = 250; // lo que paga el comprador; a Factoría le llegan 200
const COLUMNAS_PEDIDOS = ['pedido', 'fecha', 'estado', 'pago_mp', 'total', 'items', 'entrega',
  'nombre', 'email', 'telefono', 'direccion', 'localidad', 'departamento', 'nota', 'avisado'];

function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k) || ''; }
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function hoja_(nombre) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre); }

function leerStock_() {
  const filas = hoja_('Stock').getDataRange().getValues();
  const cab = filas.shift();
  return filas.filter(f => f[0]).map((f, i) => {
    const o = { fila: i + 2, stock: {} };
    cab.forEach((c, j) => { if (TALLES.includes(c)) o.stock[c] = Number(f[j]) || 0; else o[c] = f[j]; });
    o.precio = Number(o.precio);
    return o;
  });
}

function hojaPedidos_() {
  let h = hoja_('Pedidos');
  if (!h) { h = SpreadsheetApp.getActiveSpreadsheet().insertSheet('Pedidos'); h.appendRow(COLUMNAS_PEDIDOS); }
  return h;
}

// GET ?accion=stock → catálogo con stock para la página
function doGet(e) {
  if ((e.parameter || {}).accion === 'stock') {
    const productos = leerStock_().map(p => ({ id: p.id, nombre: p.nombre, desc: p.desc, img: p.img, precio: p.precio, stock: p.stock }));
    return json_({ ok: true, productos });
  }
  return json_({ ok: false });
}

// POST: o la página crea un pedido, o Mercado Pago avisa de un pago.
function doPost(e) {
  let cuerpo = {};
  try { cuerpo = JSON.parse(e.postData.contents || '{}'); } catch (err) { /* aviso sin cuerpo JSON */ }
  if (cuerpo.accion === 'pedido') return crearPedido_(cuerpo);
  return avisoMercadoPago_(e, cuerpo);
}

function crearPedido_(cuerpo) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const stock = leerStock_();
    const c = cuerpo.comprador || {};
    if (!['montevideo', 'interior'].includes(c.entrega)) return json_({ ok: false, error: 'Entrega no válida.' });
    const items = [];
    for (const it of (cuerpo.items || [])) {
      const p = stock.find(x => x.id === it.id);
      const n = Math.floor(Number(it.cantidad));
      if (!p || !TALLES.includes(it.talle) || !(n > 0)) return json_({ ok: false, error: 'Pedido no válido.' });
      if (p.stock[it.talle] < n) return json_({ ok: false, error: `${p.nombre} talle ${it.talle} se agotó.` });
      items.push({ id: p.id, title: `${p.nombre} talle ${it.talle}`, quantity: n, unit_price: p.precio, currency_id: 'UYU', talle: it.talle });
    }
    if (!items.length) return json_({ ok: false, error: 'El pedido está vacío.' });
    if (c.entrega === 'montevideo') items.push({ id: 'envio', title: 'Envío en Montevideo', quantity: 1, unit_price: ENVIO_MONTEVIDEO, currency_id: 'UYU' });

    const pedido = 'LS-' + Utilities.formatDate(new Date(), 'America/Montevideo', 'yyMMdd-HHmmss') + '-' + Math.floor(Math.random() * 900 + 100);
    const total = items.reduce((s, i) => s + i.unit_price * i.quantity, 0);
    const url = prop_('URL_TIENDA');
    const pref = {
      items: items.map(({ talle, ...i }) => i),
      payer: { name: c.nombre, email: c.email },
      external_reference: pedido,
      notification_url: ScriptApp.getService().getUrl(),
      back_urls: { success: url + '?estado=ok', pending: url + '?estado=pendiente', failure: url + '?estado=error' },
      auto_return: 'approved',
      statement_descriptor: 'LA SAGA'
    };
    const r = UrlFetchApp.fetch('https://api.mercadopago.com/checkout/preferences', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { Authorization: 'Bearer ' + prop_('MP_ACCESS_TOKEN') }, payload: JSON.stringify(pref)
    });
    const mp = JSON.parse(r.getContentText());
    if (r.getResponseCode() >= 300 || !mp.init_point) return json_({ ok: false, error: 'Mercado Pago no respondió.' });

    const fila = { pedido, fecha: new Date(), estado: 'esperando pago', pago_mp: '', total,
      items: JSON.stringify(items.filter(i => i.id !== 'envio').map(i => ({ id: i.id, talle: i.talle, cantidad: i.quantity, titulo: i.title }))),
      entrega: c.entrega, nombre: c.nombre, email: c.email, telefono: c.telefono, direccion: c.direccion || '',
      localidad: c.localidad || '', departamento: c.departamento || '', nota: c.nota || '', avisado: '' };
    hojaPedidos_().appendRow(COLUMNAS_PEDIDOS.map(k => fila[k]));
    return json_({ ok: true, url: mp.init_point });
  } finally { lock.releaseLock(); }
}

// Mercado Pago avisa con ?type=payment&data.id=… (o topic/id). Apps Script no ve los encabezados,
// así que no confiamos en el aviso: siempre se consulta el pago a la API con el token.
function avisoMercadoPago_(e, cuerpo) {
  const q = e.parameter || {};
  const tipo = q.type || q.topic || cuerpo.type;
  const id = q['data.id'] || q.id || (cuerpo.data && cuerpo.data.id);
  if (tipo !== 'payment' || !id) return json_({ ok: true });

  const r = UrlFetchApp.fetch('https://api.mercadopago.com/v1/payments/' + encodeURIComponent(id), {
    muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + prop_('MP_ACCESS_TOKEN') }
  });
  if (r.getResponseCode() !== 200) return json_({ ok: true });
  const pago = JSON.parse(r.getContentText());

  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const h = hojaPedidos_();
    const datos = h.getDataRange().getValues();
    const col = k => COLUMNAS_PEDIDOS.indexOf(k);
    const i = datos.findIndex(f => f[col('pedido')] === pago.external_reference);
    if (i < 1) return json_({ ok: true });
    const fila = datos[i];
    h.getRange(i + 1, col('estado') + 1).setValue(pago.status);
    h.getRange(i + 1, col('pago_mp') + 1).setValue(String(pago.id));
    if (pago.status !== 'approved' || fila[col('avisado')]) return json_({ ok: true });

    const pedido = Object.fromEntries(COLUMNAS_PEDIDOS.map((k, j) => [k, fila[j]]));
    pedido.items = JSON.parse(pedido.items);
    pedido.pago_mp = pago.id;
    const faltante = descontarStock_(pedido.items);
    avisarFactoria_(pedido, faltante);
    avisarComprador_(pedido);
    h.getRange(i + 1, col('avisado') + 1).setValue(new Date());
  } finally { lock.releaseLock(); }
  return json_({ ok: true });
}

// Devuelve lo que no alcanzó (dos compras simultáneas del último talle) para que se vea en el mail.
function descontarStock_(items) {
  const h = hoja_('Stock');
  const cab = h.getRange(1, 1, 1, h.getLastColumn()).getValues()[0];
  const stock = leerStock_();
  const faltante = [];
  items.forEach(it => {
    const p = stock.find(x => x.id === it.id); if (!p) return;
    const c = cab.indexOf(it.talle) + 1;
    const queda = p.stock[it.talle] - it.cantidad;
    if (queda < 0) faltante.push(`${it.titulo}: faltan ${-queda}`);
    p.stock[it.talle] = Math.max(queda, 0);
    h.getRange(p.fila, c).setValue(p.stock[it.talle]);
  });
  return faltante;
}

function textoEntrega_(p) {
  if (p.entrega === 'montevideo') return `Envío en Montevideo: ${p.direccion}, ${p.localidad}`;
  return `DAC al interior: ${p.direccion}, ${p.localidad}, ${p.departamento} (paga el envío al recibir)`;
}

function avisarFactoria_(p, faltante) {
  const lineas = p.items.map(i => `- ${i.titulo} × ${i.cantidad}`).join('\n');
  const cuerpo = [
    `Nuevo pedido pago de la tienda de La Saga (${p.pedido}).`, '',
    lineas, '',
    `Entrega: ${textoEntrega_(p)}`,
    `Nombre: ${p.nombre}`, `Celular: ${p.telefono}`, `Mail: ${p.email}`,
    p.nota ? `Nota: ${p.nota}` : '',
    faltante.length ? `\nATENCIÓN, no alcanzó el stock: ${faltante.join('; ')}` : '', '',
    `Pago de Mercado Pago n.º ${p.pago_mp}, aprobado.`
  ].filter(x => x !== null).join('\n');
  MailApp.sendEmail({
    to: prop_('MAIL_FACTORIA'), cc: prop_('MAIL_BANDA'), replyTo: p.email,
    subject: `Pedido La Saga ${p.pedido}: ${p.items.reduce((s, i) => s + i.cantidad, 0)} remera(s)`,
    body: cuerpo, name: 'Tienda La Saga'
  });
}

function avisarComprador_(p) {
  const lineas = p.items.map(i => `- ${i.titulo} × ${i.cantidad}`).join('\n');
  MailApp.sendEmail({
    to: p.email, replyTo: prop_('MAIL_BANDA'),
    subject: 'Tu compra en la tienda de La Saga',
    body: `Hola, ${p.nombre}:\n\nRecibimos tu pago. Esto es lo que pediste:\n\n${lineas}\n\nEntrega: ${textoEntrega_(p)}\n\nTe escribimos para coordinar. Gracias por bancarnos.\n\nLa Saga`,
    name: 'La Saga'
  });
}
