/**
 * BACKEND DE PEDIDOS MANGO
 * Google Sheets + Apps Script + GitHub Pages
 */
const HOJAS = Object.freeze({
  PRODUCTOS: 'Productos',
  CLIENTES: 'Clientes',
  PEDIDOS: 'Pedidos',
  CONFIG: 'Config',
  RESUMEN: 'Resumen'
});

function ss() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Control de pedidos')
    .addItem('Abrir recepción de pedidos', 'abrirVentanaDePedidos')
    .addItem('Cerrar y consolidar pedidos', 'cerrarVentanaDePedidos')
    .addItem('Actualizar resumen', 'generarResumenPedidos')
    .addToUi();
}

function doGet(e) {
  try {
    const accion = String(e && e.parameter && e.parameter.action || '').toLowerCase();
    if (accion === 'catalogo') return jsonResponse(obtenerCatalogo());
    if (accion === 'cliente') return jsonResponse(buscarCliente(e.parameter.telefono));
    if (accion === 'estado') return jsonResponse(estadoPedido(e.parameter.telefono));
    if (accion === 'config') return jsonResponse(obtenerConfiguracionPublica());
    return jsonResponse({ error: 'Acción no reconocida.' });
  } catch (error) {
    console.error(error);
    return jsonResponse({ error: error.message || 'Error interno del servidor.' });
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) throw new Error('Solicitud vacía.');
    const body = JSON.parse(e.postData.contents);
    if (body.action === 'guardarPedido') return jsonResponse(guardarPedido(body));
    return jsonResponse({ error: 'Acción no reconocida.' });
  } catch (error) {
    console.error(error);
    return jsonResponse({ error: error.message || 'No se pudo procesar el pedido.' });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function obtenerHoja(nombre) {
  const hoja = ss().getSheetByName(nombre);
  if (!hoja) throw new Error('No existe la hoja ' + nombre + '.');
  return hoja;
}

function objetoDesdeFila(encabezados, fila) {
  const obj = {};
  encabezados.forEach((clave, i) => obj[clave] = fila[i]);
  return obj;
}

function normalizarTelefono(valor) {
  let telefono = String(valor || '').replace(/\D/g, '');
  if (telefono.length === 11 && telefono.indexOf('506') === 0) telefono = telefono.slice(3);
  if (telefono.length !== 8) throw new Error('El teléfono debe tener 8 dígitos.');
  return telefono;
}

function textoLimpio(valor, maximo) {
  return String(valor || '').replace(/\s+/g, ' ').trim().slice(0, maximo);
}

function obtenerCatalogo() {
  const hoja = obtenerHoja(HOJAS.PRODUCTOS);
  const datos = hoja.getDataRange().getValues();
  if (datos.length < 2) return [];
  const encabezados = datos[0];
  return datos.slice(1)
    .map(fila => objetoDesdeFila(encabezados, fila))
    .filter(producto => /^s[ií]$/i.test(String(producto.Disponible).trim()))
    .map(producto => ({
      ID_Producto: String(producto.ID_Producto),
      Nombre: String(producto.Nombre),
      Precio: Number(producto.Precio),
      Unidad: String(producto.Unidad || ''),
      Foto: String(producto.Foto || '')
    }));
}

function buscarCliente(telefonoEntrada) {
  const telefono = normalizarTelefono(telefonoEntrada);
  const hoja = obtenerHoja(HOJAS.CLIENTES);
  const datos = hoja.getDataRange().getValues();
  if (datos.length < 2) return null;
  const encabezados = datos[0];
  const fila = datos.slice(1).find(registro => normalizarTelefonoSeguro_(registro[0]) === telefono);
  return fila ? objetoDesdeFila(encabezados, fila) : null;
}

function normalizarTelefonoSeguro_(valor) {
  let telefono = String(valor || '').replace(/\D/g, '');
  if (telefono.length === 11 && telefono.indexOf('506') === 0) telefono = telefono.slice(3);
  return telefono;
}

function guardarOActualizarCliente(telefono, nombre, direccion, gps) {
  const hoja = obtenerHoja(HOJAS.CLIENTES);
  const datos = hoja.getDataRange().getValues();
  const indice = datos.findIndex((fila, i) => i > 0 && normalizarTelefonoSeguro_(fila[0]) === telefono);
  if (indice === -1) {
    hoja.appendRow([telefono, nombre, direccion, gps, '']);
    return;
  }
  const existente = datos[indice];
  hoja.getRange(indice + 1, 2, 1, 3).setValues([[
    nombre || existente[1],
    direccion || existente[2],
    gps || existente[3]
  ]]);
}

function guardarPedido(body) {
  const config = leerConfig_();
  if (/^no$/i.test(String(config.PedidosAbiertos || 'Si').trim())) {
    throw new Error('La recepción de pedidos está cerrada por el momento.');
  }

  const telefono = normalizarTelefono(body.telefono);
  const nombre = textoLimpio(body.nombre, 80);
  const direccion = textoLimpio(body.direccion, 300);
  const gps = validarGps_(body.gps);
  if (!nombre || !direccion) throw new Error('Nombre y dirección son obligatorios.');
  if (!Array.isArray(body.items) || !body.items.length) throw new Error('El carrito está vacío.');

  const catalogo = obtenerCatalogo();
  const productos = new Map(catalogo.map(producto => [String(producto.ID_Producto), producto]));
  const items = body.items.map(item => {
    const producto = productos.get(String(item.ID_Producto));
    const cantidad = Math.floor(Number(item.Cantidad));
    if (!producto) throw new Error('Uno de los productos ya no está disponible.');
    if (!Number.isFinite(cantidad) || cantidad < 1 || cantidad > 99) throw new Error('Cantidad de producto inválida.');
    return { producto, cantidad };
  });

  const total = items.reduce((suma, item) => suma + item.cantidad * item.producto.Precio, 0);
  const montoMinimo = Number(config.MontoMinimo || 0);
  if (montoMinimo > 0 && total < montoMinimo) {
    throw new Error('El monto mínimo del pedido es ₡' + montoMinimo.toLocaleString('es-CR') + '.');
  }

  const bloqueo = LockService.getScriptLock();
  bloqueo.waitLock(20000);
  let idPedido;
  try {
    guardarOActualizarCliente(telefono, nombre, direccion, gps);
    const hoja = obtenerHoja(HOJAS.PEDIDOS);
    idPedido = Utilities.getUuid();
    const fecha = new Date();
    const filas = items.map(item => [
      idPedido,
      telefono,
      item.producto.ID_Producto,
      item.producto.Nombre,
      item.cantidad,
      item.producto.Precio,
      fecha,
      'Recibido'
    ]);
    hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, 8).setValues(filas);
  } finally {
    bloqueo.releaseLock();
  }

  notificarPedido_(config, { idPedido, telefono, nombre, direccion, items, total });
  return { ok: true, idPedido, total };
}

function validarGps_(valor) {
  const gps = String(valor || '').trim();
  if (!gps) return '';
  const partes = gps.split(',').map(Number);
  if (partes.length !== 2 || !partes.every(Number.isFinite)) return '';
  if (partes[0] < -90 || partes[0] > 90 || partes[1] < -180 || partes[1] > 180) return '';
  return partes[0].toFixed(6) + ',' + partes[1].toFixed(6);
}

function estadoPedido(telefonoEntrada) {
  const telefono = normalizarTelefono(telefonoEntrada);
  const hoja = obtenerHoja(HOJAS.PEDIDOS);
  const datos = hoja.getDataRange().getValues();
  if (datos.length < 2) return [];
  const encabezados = datos[0];
  return datos.slice(1)
    .filter(fila => normalizarTelefonoSeguro_(fila[1]) === telefono)
    .map(fila => objetoDesdeFila(encabezados, fila))
    .sort((a, b) => new Date(a.Fecha) - new Date(b.Fecha));
}

function leerConfig_() {
  const hoja = obtenerHoja(HOJAS.CONFIG);
  const filas = hoja.getDataRange().getValues();
  const config = {};
  filas.slice(1).forEach(fila => {
    const clave = String(fila[0] || '').trim();
    if (clave) config[clave] = fila[1];
  });
  return config;
}

function obtenerConfiguracionPublica() {
  const config = leerConfig_();
  return {
    PedidosAbiertos: !/^no$/i.test(String(config.PedidosAbiertos || 'Si').trim()),
    FechaCorte: config.FechaCorte || '',
    MontoMinimo: Number(config.MontoMinimo || 0),
    WhatsAppNegocio: String(config.WhatsAppNegocio || '').replace(/\D/g, '')
  };
}

function escribirConfig_(clave, valor) {
  const hoja = obtenerHoja(HOJAS.CONFIG);
  const datos = hoja.getDataRange().getValues();
  const indice = datos.findIndex((fila, i) => i > 0 && String(fila[0]).trim() === clave);
  if (indice === -1) hoja.appendRow([clave, valor]);
  else hoja.getRange(indice + 1, 2).setValue(valor);
}

function abrirVentanaDePedidos() {
  escribirConfig_('PedidosAbiertos', 'Si');
  SpreadsheetApp.getActive().toast('La recepción de pedidos está abierta.', 'Control de pedidos', 5);
}

function cerrarVentanaDePedidos() {
  const hoja = obtenerHoja(HOJAS.PEDIDOS);
  const datos = hoja.getDataRange().getValues();
  if (datos.length > 1) {
    const estados = datos.slice(1).map(fila => [fila[7] === 'Recibido' ? 'Cerrado' : fila[7]]);
    hoja.getRange(2, 8, estados.length, 1).setValues(estados);
  }
  escribirConfig_('PedidosAbiertos', 'No');
  generarResumenPedidos();
  SpreadsheetApp.getActive().toast('Pedidos cerrados y resumen actualizado.', 'Control de pedidos', 6);
}

function consolidarPedidos() {
  const hoja = obtenerHoja(HOJAS.PEDIDOS);
  const datos = hoja.getDataRange().getValues();
  if (datos.length < 2) return {};
  const encabezados = datos[0];
  const totales = {};
  datos.slice(1).forEach(fila => {
    const pedido = objetoDesdeFila(encabezados, fila);
    if (pedido.Estado === 'Recibido' || pedido.Estado === 'Cerrado') {
      const clave = String(pedido.NombreProducto);
      totales[clave] = (totales[clave] || 0) + Number(pedido.Cantidad || 0);
    }
  });
  return totales;
}

function generarResumenPedidos() {
  const libro = ss();
  const totales = consolidarPedidos();
  let hoja = libro.getSheetByName(HOJAS.RESUMEN);
  if (!hoja) hoja = libro.insertSheet(HOJAS.RESUMEN);
  hoja.clearContents();
  const filas = [['Producto', 'Cantidad total']]
    .concat(Object.keys(totales).sort().map(nombre => [nombre, totales[nombre]]));
  hoja.getRange(1, 1, filas.length, 2).setValues(filas);
  hoja.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#1F4B3F').setFontColor('#FFFFFF');
  hoja.setFrozenRows(1);
  hoja.autoResizeColumns(1, 2);
  return totales;
}

function notificarPedido_(config, pedido) {
  const email = String(config.EmailNotificacion || '').trim();
  if (!email || !/^\S+@\S+\.\S+$/.test(email)) return;
  const detalle = pedido.items.map(item =>
    '- ' + item.cantidad + ' × ' + item.producto.Nombre + ' = ₡' +
    (item.cantidad * item.producto.Precio).toLocaleString('es-CR')
  ).join('\n');
  const mensaje = [
    'Nuevo pedido recibido',
    '',
    'Código: ' + pedido.idPedido,
    'Cliente: ' + pedido.nombre,
    'WhatsApp: ' + pedido.telefono,
    'Dirección: ' + pedido.direccion,
    '',
    detalle,
    '',
    'Total: ₡' + pedido.total.toLocaleString('es-CR')
  ].join('\n');
  try {
    MailApp.sendEmail(email, 'Nuevo pedido - ' + pedido.nombre, mensaje);
  } catch (error) {
    console.error('No se pudo enviar la notificación: ' + error.message);
  }
}

function calcularRutaOptima(origenLatLng) {
  const origen = validarGps_(origenLatLng);
  if (!origen) return { error: 'El origen no tiene un GPS válido.' };
  const hojaPedidos = obtenerHoja(HOJAS.PEDIDOS);
  const datosPedidos = hojaPedidos.getDataRange().getValues();
  const encabezadosPedidos = datosPedidos[0];
  const telefonos = [...new Set(datosPedidos.slice(1)
    .map(fila => objetoDesdeFila(encabezadosPedidos, fila))
    .filter(pedido => pedido.Estado === 'Cerrado')
    .map(pedido => normalizarTelefonoSeguro_(pedido.Telefono)))];

  const hojaClientes = obtenerHoja(HOJAS.CLIENTES);
  const datosClientes = hojaClientes.getDataRange().getValues();
  const encabezadosClientes = datosClientes[0];
  const paradas = telefonos.map(telefono => {
    const fila = datosClientes.slice(1).find(registro => normalizarTelefonoSeguro_(registro[0]) === telefono);
    return fila ? validarGps_(objetoDesdeFila(encabezadosClientes, fila).Gps) : '';
  }).filter(Boolean);

  if (!paradas.length) return { error: 'No hay pedidos cerrados con GPS registrado.' };
  if (paradas.length > 25) return { error: 'La ruta supera el máximo de 25 paradas y debe dividirse.' };
  const apiKey = PropertiesService.getScriptProperties().getProperty('MAPS_API_KEY');
  if (!apiKey) return { error: 'Falta configurar MAPS_API_KEY.' };
  const url = 'https://maps.googleapis.com/maps/api/directions/json' +
    '?origin=' + encodeURIComponent(origen) +
    '&destination=' + encodeURIComponent(origen) +
    '&waypoints=' + encodeURIComponent('optimize:true|' + paradas.join('|')) +
    '&key=' + encodeURIComponent(apiKey);
  const respuesta = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  return JSON.parse(respuesta.getContentText());
}
