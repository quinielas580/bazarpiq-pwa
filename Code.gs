// ════════════════════════════════════════════════════════
//  BAZAR-PIQ  —  Code.gs  (versión completa v4)
//  Admin WA: 524423774943
// ════════════════════════════════════════════════════════

var ADMIN_PASS   = "BazarPIQ2025Admin!Seguro";
var ADMIN_WA     = "524423774943";
var SALDO_REGALO = 100;   // Saldo inicial de regalo

// ── Comisiones ─────────────────────────────────────────────
function calcularComision(precio) {
  var p = parseFloat(precio);
  if (isNaN(p) || p <= 0) return 0;
  if (p <= 30)     return Math.round(p * 0.005 * 100) / 100;
  if (p <= 50)     return 2;
  if (p <= 100)    return 3;
  if (p <= 3000)   return Math.min(Math.round(p * 0.05 * 100) / 100, 150);
  if (p <= 24999)  return 150;
  if (p <= 39999)  return 300;
  if (p <= 100000) return 500;
  return 700;
}

// Comisión por publicar vacante de empleo
var COMISION_VACANTE = 5;

// ── doGet ──────────────────────────────────────────────────
function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Config global ──────────────────────────────────────────
function obtenerConfig() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("CONFIG");
  if (!hoja) {
    hoja = ss.insertSheet("CONFIG");
    hoja.appendRow(["clave","valor"]);
    hoja.appendRow(["modo_gratuito",     "NO"]);
    hoja.appendRow(["admin_pass",        ADMIN_PASS]);
    hoja.appendRow(["regalo_activo",     "NO"]);
    hoja.appendRow(["version_app",       "1.0.0"]);
    hoja.appendRow(["hay_actualizacion", "NO"]);
    hoja.appendRow(["url_descarga",      ""]);
  }
  var datos = hoja.getDataRange().getValues();
  var cfg = {};
  for (var i = 1; i < datos.length; i++) cfg[String(datos[i][0])] = String(datos[i][1]);
  return cfg;
}

function setConfig(clave, valor) {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("CONFIG") || ss.insertSheet("CONFIG");
  var datos = hoja.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][0]) === clave) { hoja.getRange(i+1,2).setValue(valor); return; }
  }
  hoja.appendRow([clave, valor]);
}

function modoGratuitoActivo() {
  return String(obtenerConfig()["modo_gratuito"]||"NO").toUpperCase()==="SI";
}

function regaloActivo() {
  return String(obtenerConfig()["regalo_activo"]||"NO").toUpperCase()==="SI";
}

// ── Hojas ──────────────────────────────────────────────────
function obtenerHoja(nombre) {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
  if (hoja.getLastRow() === 0) {
    if (nombre === "TIENDAS") {
      hoja.appendRow(["fecha","nombre_tienda","pass","tel_recuperacion","h_abre","h_cierra",
                       "estatus","ventas","dias_servicio","categoria","entrega_modo",
                       "comunidades_entrega","whatsapp","ubicacion_tienda","saldo","es_gratis_origen"]);
    } else if (nombre === "PRODUCTOS_APP") {
      hoja.appendRow(["fecha","nombre_tienda","prod_nombre","precio","desc","stock","ubicacion",
                       "whatsapp","foto","entrega","categoria","id_prod","unidad_medida",
                       "publicado_gratis","condicion"]);
    } else if (nombre === "PRODUCTOS_UNICOS") {
      hoja.appendRow(["fecha","vendedor","prod_nombre","precio","desc","whatsapp","foto",
                       "entrega","categoria","id_prod","unidad_medida","unidades_total",
                       "publicado_gratis","condicion"]);
    } else if (nombre === "CALIFICACIONES") {
      hoja.appendRow(["fecha","tienda","estrellas"]);
    } else if (nombre === "CALIFICACIONES_APP") {
      hoja.appendRow(["fecha","usuario","estrellas","comentario"]);
    } else if (nombre === "EMPRESAS") {
      // Igual que TIENDAS pero para reclutadores — col 15 = saldo, col 16 = es_gratis_origen
      hoja.appendRow(["fecha","nombre_empresa","pass","tel","h_abre","h_cierra",
                       "estatus","publicaciones","ubicacion","sector","whatsapp","saldo","es_gratis_origen"]);
    } else if (nombre === "OFERTAS_EMPLEO") {
      hoja.appendRow(["fecha","empresa","puesto","sueldo","horario","descripcion","whatsapp","ubicacion","id_oferta"]);
    } else if (nombre === "BUSCO_EMPLEO") {
      hoja.appendRow(["fecha","nombre","oficio","descripcion","whatsapp","id_solicitud"]);
    } else if (nombre === "MOVIMIENTOS_SALDO") {
      hoja.appendRow(["fecha","tienda","tipo","monto","descripcion","saldo_anterior","saldo_nuevo"]);
    }
  }
  return hoja;
}

// ══════════════════════════════════════════════════════════
//  ADMIN
// ══════════════════════════════════════════════════════════
function loginAdmin(pass) {
  var cfg = obtenerConfig();
  if (pass !== String(cfg["admin_pass"]||ADMIN_PASS)) return null;
  return { ok: true, modo_gratuito: String(cfg["modo_gratuito"]||"NO").toUpperCase()==="SI",
           regalo_activo: String(cfg["regalo_activo"]||"NO").toUpperCase()==="SI",
           version_app: String(cfg["version_app"]||"1.0.0"),
           hay_actualizacion: String(cfg["hay_actualizacion"]||"NO").toUpperCase()==="SI",
           url_descarga: String(cfg["url_descarga"]||"") };
}

function cambiarPassAdmin(actual, nueva, confirm) {
  var cfg = obtenerConfig();
  if (actual !== String(cfg["admin_pass"]||ADMIN_PASS)) return {ok:false,msg:"Contraseña actual incorrecta"};
  if (nueva !== confirm)  return {ok:false,msg:"Las nuevas contraseñas no coinciden"};
  if (nueva.length < 10)  return {ok:false,msg:"Mínimo 10 caracteres"};
  setConfig("admin_pass", nueva);
  return {ok:true};
}

function toggleModoGratuito(activar) {
  setConfig("modo_gratuito", activar?"SI":"NO");
  return {ok:true, modo_gratuito:activar};
}

function toggleRegalo(activar) {
  setConfig("regalo_activo", activar?"SI":"NO");
  return {ok:true, regalo_activo:activar};
}

function obtenerEstadoConfig() {
  var cfg = obtenerConfig();
  return {
    modo_gratuito:    String(cfg["modo_gratuito"]||"NO").toUpperCase()==="SI",
    regalo_activo:    String(cfg["regalo_activo"]||"NO").toUpperCase()==="SI",
    version_app:      String(cfg["version_app"]||"1.0.0"),
    hay_actualizacion:String(cfg["hay_actualizacion"]||"NO").toUpperCase()==="SI",
    url_descarga:     String(cfg["url_descarga"]||"")
  };
}

// Subir nueva versión (admin guarda datos de actualización)
function publicarActualizacion(version, urlDescarga, notas) {
  setConfig("version_app",       version);
  setConfig("hay_actualizacion", "SI");
  setConfig("url_descarga",      urlDescarga);
  setConfig("notas_update",      notas||"");
  return {ok:true};
}

// El usuario "descargó" la actualización — quita el punto rojo
function marcarActualizacionVista() {
  // No cambiamos hay_actualizacion global, solo registramos en local del cliente
  return {ok:true};
}

function adminObtenerResumen() {
  var datos = obtenerHoja("TIENDAS").getDataRange().getValues();
  var lista = [];
  for (var i = 1; i < datos.length; i++) {
    lista.push({ nombre:String(datos[i][1]||""), estatus:String(datos[i][6]||"ABIERTO"),
                 ventas:datos[i][7]||0, saldo:parseFloat(datos[i][14]||0),
                 es_gratis_origen:String(datos[i][15]||"NO") });
  }
  return lista;
}

function adminRecargaSaldo(nombre, monto) {
  var hoja = obtenerHoja("TIENDAS"), datos = hoja.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][1]).toLowerCase()===String(nombre).toLowerCase()) {
      var ant = parseFloat(datos[i][14]||0), nvo = ant+parseFloat(monto);
      hoja.getRange(i+1,15).setValue(nvo);
      _registrarMovimiento(nombre,"RECARGA",monto,"Recarga por administrador",ant,nvo);
      return {ok:true,saldo:nvo};
    }
  }
  return {ok:false};
}

function adminDescontarSaldo(nombre, monto, motivo) {
  var hoja = obtenerHoja("TIENDAS"), datos = hoja.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][1]).toLowerCase()===String(nombre).toLowerCase()) {
      var ant = parseFloat(datos[i][14]||0), nvo = Math.max(0,ant-parseFloat(monto));
      hoja.getRange(i+1,15).setValue(nvo);
      _registrarMovimiento(nombre,"DESCUENTO",monto,motivo||"Ajuste manual",ant,nvo);
      return {ok:true,saldo:nvo};
    }
  }
  return {ok:false};
}

function adminObtenerMovimientos(nombre) {
  var datos = obtenerHoja("MOVIMIENTOS_SALDO").getDataRange().getValues(), lista=[];
  for (var i = 1; i < datos.length; i++) {
    if (String(datos[i][1]).toLowerCase()===String(nombre).toLowerCase()) {
      lista.push({ fecha:datos[i][0]?Utilities.formatDate(new Date(datos[i][0]),Session.getScriptTimeZone(),"dd/MM/yyyy HH:mm"):"",
                   tipo:String(datos[i][2]||""), monto:parseFloat(datos[i][3]||0),
                   desc:String(datos[i][4]||""), saldo_nuevo:parseFloat(datos[i][6]||0) });
    }
  }
  return lista.reverse();
}

function _registrarMovimiento(tienda, tipo, monto, desc, ant, nvo) {
  obtenerHoja("MOVIMIENTOS_SALDO").appendRow([new Date(),tienda,tipo,monto,desc,ant,nvo]);
}

// Calificaciones de la app (solo admin puede ver comentarios)
function guardarCalificacionApp(usuario, estrellas, comentario) {
  obtenerHoja("CALIFICACIONES_APP").appendRow([new Date(),usuario,estrellas,comentario||""]);
  return "OK";
}

function adminObtenerCalificacionesApp() {
  var datos = obtenerHoja("CALIFICACIONES_APP").getDataRange().getValues();
  if (datos.length<2) return {promedio:0, total:0, comentarios:[]};
  var suma=0, lista=[];
  for (var i=1;i<datos.length;i++) {
    suma += parseFloat(datos[i][2]||0);
    lista.push({fecha:datos[i][0]?Utilities.formatDate(new Date(datos[i][0]),Session.getScriptTimeZone(),"dd/MM/yyyy"):"",
                usuario:String(datos[i][1]||"Anónimo"), estrellas:parseFloat(datos[i][2]||0),
                comentario:String(datos[i][3]||"")});
  }
  return {promedio:((suma/(datos.length-1)).toFixed(1)), total:datos.length-1, comentarios:lista.reverse()};
}

function obtenerPromedioApp() {
  var datos = obtenerHoja("CALIFICACIONES_APP").getDataRange().getValues();
  if (datos.length<2) return {promedio:0, total:0};
  var suma=0;
  for (var i=1;i<datos.length;i++) suma+=parseFloat(datos[i][2]||0);
  return {promedio:((suma/(datos.length-1)).toFixed(1)), total:datos.length-1};
}

// Limpiar fotos de Drive de productos eliminados / sin stock
function limpiarFotosHuerfanas() {
  var hojaP = obtenerHoja("PRODUCTOS_APP");
  var datos  = hojaP.getDataRange().getValues();
  var fotoIds = {};
  for (var i = 1; i < datos.length; i++) {
    var url = String(datos[i][8]||"");
    var match = url.match(/id=([^&]+)/);
    if (match) fotoIds[match[1]] = true;
  }
  // Buscar archivos en la carpeta y eliminar huérfanos
  try {
    var folder = DriveApp.getFolderById("16XRhzFqFf6F3qlebw9NFfctbiH4yCXPV");
    var files  = folder.getFiles();
    var borrados = 0;
    while (files.hasNext()) {
      var file = files.next();
      if (!fotoIds[file.getId()]) {
        file.setTrashed(true);
        borrados++;
      }
    }
    return {ok:true, borrados:borrados};
  } catch(e) {
    return {ok:false, msg:e.toString()};
  }
}

// ══════════════════════════════════════════════════════════
//  TIENDAS
// ══════════════════════════════════════════════════════════
function registrarTiendaServidor(obj) {
  var hoja = obtenerHoja("TIENDAS"), datos = hoja.getDataRange().getValues();
  for (var i=1;i<datos.length;i++) {
    if (datos[i][1].toString().toLowerCase()===obj.nombre.toLowerCase()) return "EXISTE";
  }
  var esGratis   = modoGratuitoActivo()?"SI":"NO";
  var saldoInicial = regaloActivo() ? SALDO_REGALO : 0;
  hoja.appendRow([new Date(),obj.nombre,obj.pass,obj.tel,obj.h_abre,obj.h_cierra,
                  "ABIERTO",0,obj.dias,obj.categoria,obj.entrega_modo,
                  obj.comunidades_entrega,obj.whatsapp,obj.ubicacion_tienda,
                  saldoInicial, esGratis]);
  if (saldoInicial > 0) {
    _registrarMovimiento(obj.nombre,"REGALO",saldoInicial,"Saldo de bienvenida (regalo de inicio)",0,saldoInicial);
  }
  return "OK";
}

function loginTienda(nombre, pass) {
  var datos = obtenerHoja("TIENDAS").getDataRange().getValues();
  for (var i=1;i<datos.length;i++) {
    if (String(datos[i][1]).toLowerCase()===String(nombre).toLowerCase() &&
        String(datos[i][2])===String(pass)) {
      var cfg = obtenerConfig();
      return {
        nombre:              String(datos[i][1]),
        tel:                 String(datos[i][3]),
        estatus:             String(datos[i][6]||"ABIERTO"),
        ventas:              datos[i][7]||0,
        dias:                String(datos[i][8]||""),
        categoria:           String(datos[i][9]||""),
        entrega_modo:        String(datos[i][10]||"Punto Medio"),
        comunidades_entrega: String(datos[i][11]||""),
        whatsapp:            String(datos[i][12]||datos[i][3]),
        ubicacion_tienda:    String(datos[i][13]||""),
        saldo:               parseFloat(datos[i][14]||0),
        es_gratis_origen:    String(datos[i][15]||"NO"),
        modo_gratuito_global:String(cfg["modo_gratuito"]||"NO").toUpperCase()==="SI",
        hay_actualizacion:   String(cfg["hay_actualizacion"]||"NO").toUpperCase()==="SI",
        version_app:         String(cfg["version_app"]||"1.0.0"),
        url_descarga:        String(cfg["url_descarga"]||"")
      };
    }
  }
  return null;
}

function cambiarPassTienda(nombre, actual, nueva, confirm) {
  var hoja = obtenerHoja("TIENDAS"), datos = hoja.getDataRange().getValues();
  for (var i=1;i<datos.length;i++) {
    if (String(datos[i][1]).toLowerCase()===String(nombre).toLowerCase()) {
      if (String(datos[i][2])!==String(actual)) return {ok:false,msg:"Contraseña actual incorrecta"};
      if (nueva!==confirm) return {ok:false,msg:"Las nuevas contraseñas no coinciden"};
      if (nueva.length<8)  return {ok:false,msg:"Mínimo 8 caracteres"};
      hoja.getRange(i+1,3).setValue(nueva);
      return {ok:true};
    }
  }
  return {ok:false,msg:"Tienda no encontrada"};
}

function actualizarTienda(obj) {
  var hoja = obtenerHoja("TIENDAS"), datos = hoja.getDataRange().getValues();
  for (var i=1;i<datos.length;i++) {
    if (String(datos[i][1]).toLowerCase()===String(obj.nombre).toLowerCase()) {
      if (obj.whatsapp)         hoja.getRange(i+1,13).setValue(obj.whatsapp);
      if (obj.h_abre)           hoja.getRange(i+1, 5).setValue(obj.h_abre);
      if (obj.h_cierra)         hoja.getRange(i+1, 6).setValue(obj.h_cierra);
      if (obj.ubicacion_tienda) hoja.getRange(i+1,14).setValue(obj.ubicacion_tienda);
      if (obj.entrega_modo)     hoja.getRange(i+1,11).setValue(obj.entrega_modo);
      if (obj.comunidades_entrega!==undefined) hoja.getRange(i+1,12).setValue(obj.comunidades_entrega);
      return "OK";
    }
  }
  return "NO_ENCONTRADO";
}

function toggleEstatus(nombre, nuevo) {
  var hoja = obtenerHoja("TIENDAS"), datos = hoja.getDataRange().getValues();
  for (var i=1;i<datos.length;i++) {
    if (String(datos[i][1]).toLowerCase()===String(nombre).toLowerCase()) {
      hoja.getRange(i+1,7).setValue(nuevo); return nuevo;
    }
  }
  return "NO_ENCONTRADO";
}

function obtenerInfoTienda(nombre) {
  var datos = obtenerHoja("TIENDAS").getDataRange().getValues();
  for (var i=1;i<datos.length;i++) {
    if (String(datos[i][1]).toLowerCase()===String(nombre).toLowerCase()) {
      return { nombre:String(datos[i][1]), h_abre:String(datos[i][4]||""),
               h_cierra:String(datos[i][5]||""), estatus:String(datos[i][6]||"ABIERTO"),
               dias:String(datos[i][8]||""), categoria:String(datos[i][9]||""),
               entrega_modo:String(datos[i][10]||"Punto Medio"),
               comunidades_entrega:String(datos[i][11]||""),
               whatsapp:String(datos[i][12]||datos[i][3]),
               ubicacion_tienda:String(datos[i][13]||""),
               saldo:parseFloat(datos[i][14]||0),
               es_gratis_origen:String(datos[i][15]||"NO") };
    }
  }
  return null;
}

function obtenerSaldoTienda(nombre) {
  var datos = obtenerHoja("TIENDAS").getDataRange().getValues();
  for (var i=1;i<datos.length;i++) {
    if (String(datos[i][1]).toLowerCase()===String(nombre).toLowerCase()) return parseFloat(datos[i][14]||0);
  }
  return 0;
}

function obtenerMovimientosTienda(nombre) { return adminObtenerMovimientos(nombre); }

// ══════════════════════════════════════════════════════════
//  PRODUCTOS
// ══════════════════════════════════════════════════════════
function guardarProducto(obj) {
  var hoja    = obtenerHoja("PRODUCTOS_APP");
  var id      = "ID_"+new Date().getTime();
  var esGratis = (modoGratuitoActivo()||obj.es_gratis_origen==="SI")?"SI":"NO";
  hoja.appendRow([new Date(),obj.nombre_tienda,obj.prod_nombre,obj.precio,obj.desc,
                  obj.stock,obj.ubicacion,obj.whatsapp,obj.foto,obj.entrega,
                  obj.categoria,id,obj.unidad_medida||"Pieza",esGratis,obj.condicion||"Nuevo"]);
  return "OK";
}

function guardarProductoUnico(obj) {
  var hoja    = obtenerHoja("PRODUCTOS_UNICOS");
  var id      = "UNQ_"+new Date().getTime();
  var esGratis = modoGratuitoActivo()?"SI":"NO";
  hoja.appendRow([new Date(),obj.vendedor,obj.prod_nombre,obj.precio,obj.desc,
                  obj.whatsapp,obj.foto,obj.entrega,obj.categoria,id,
                  obj.unidad_medida||"Pieza",obj.unidades_total||1,esGratis,obj.condicion||"Nuevo"]);
  return "OK";
}

function actualizarProducto(id, campo, valor) {
  var hoja  = obtenerHoja("PRODUCTOS_APP");
  var datos = hoja.getDataRange().getValues();
  var colMap = {precio:4, stock:6, desc:5};
  var col = colMap[campo];
  if (!col) return "CAMPO_INVALIDO";
  for (var i=1;i<datos.length;i++) {
    if (datos[i][11]===id) { hoja.getRange(i+1,col).setValue(valor); return "OK"; }
  }
  return "NO_ENCONTRADO";
}

function actualizarPrecioServidor(id, precio) {
  return actualizarProducto(id,"precio",precio);
}

function obtenerProductos() {
  var tiendas  = obtenerHoja("TIENDAS").getDataRange().getValues();
  var dT       = obtenerHoja("PRODUCTOS_APP").getDataRange().getValues();
  var dU       = obtenerHoja("PRODUCTOS_UNICOS").getDataRange().getValues();
  var abiertas = {};
  for (var i=1;i<tiendas.length;i++) {
    if (String(tiendas[i][6])==="ABIERTO") abiertas[tiendas[i][1].toString().toLowerCase()]=true;
  }
  var lista=[];
  if (dT.length>1) dT.slice(1).forEach(function(f){
    if (abiertas[f[1].toString().toLowerCase()]) {
      lista.push({nombre_tienda:String(f[1]),prod_nombre:String(f[2]),precio:f[3],
                  desc:String(f[4]||""),stock:f[5],ubicacion:String(f[6]||""),
                  whatsapp:String(f[7]||""),foto:String(f[8]||""),entrega:String(f[9]||""),
                  categoria:String(f[10]||""),id_prod:String(f[11]||""),
                  unidad_medida:String(f[12]||"Pieza"),publicado_gratis:String(f[13]||"NO"),
                  condicion:String(f[14]||"Nuevo")});
    }
  });
  if (dU.length>1) dU.slice(1).forEach(function(f){
    lista.push({nombre_tienda:"Particular: "+String(f[1]||"Vendedor"),prod_nombre:String(f[2]||""),
                precio:f[3],desc:String(f[4]||""),whatsapp:String(f[5]||""),foto:String(f[6]||""),
                entrega:String(f[7]||""),categoria:String(f[8]||""),id_prod:String(f[9]||""),
                ubicacion:"Venta Particular",unidad_medida:String(f[10]||"Pieza"),
                unidades_total:f[11]||1,publicado_gratis:String(f[12]||"NO"),
                condicion:String(f[13]||"Nuevo")});
  });
  return lista.reverse();
}

function obtenerTodosLosProductos() {
  var datos = obtenerHoja("PRODUCTOS_APP").getDataRange().getValues();
  if (datos.length<2) return [];
  return datos.slice(1).map(function(f){
    return {nombre_tienda:String(f[1]||""),prod_nombre:String(f[2]||""),precio:f[3],
            desc:String(f[4]||""),stock:f[5],ubicacion:String(f[6]||""),
            whatsapp:String(f[7]||""),foto:String(f[8]||""),entrega:String(f[9]||""),
            categoria:String(f[10]||""),id_prod:String(f[11]||""),
            unidad_medida:String(f[12]||"Pieza"),publicado_gratis:String(f[13]||"NO"),
            condicion:String(f[14]||"Nuevo")};
  }).reverse();
}

function eliminarProducto(id) {
  var hoja = obtenerHoja("PRODUCTOS_APP"), datos = hoja.getDataRange().getValues();
  for (var i=1;i<datos.length;i++) {
    if (datos[i][11]===id) {
      // Intentar borrar foto de Drive
      var url = String(datos[i][8]||"");
      var match = url.match(/id=([^&]+)/);
      if (match) { try { DriveApp.getFileById(match[1]).setTrashed(true); } catch(e) {} }
      hoja.deleteRow(i+1);
      return "OK";
    }
  }
  return "NO_ENCONTRADO";
}

function subirFoto(base64, nombre) {
  var folderId = "16XRhzFqFf6F3qlebw9NFfctbiH4yCXPV";
  var blob     = Utilities.newBlob(Utilities.base64Decode(base64.split(",")[1]),"image/jpeg",nombre);
  var archivo  = DriveApp.getFolderById(folderId).createFile(blob);
  archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);
  return "https://drive.google.com/thumbnail?id="+archivo.getId()+"&sz=w800";
}

function registrarVenta(nombre, precio, fueGratis) {
  var hoja = obtenerHoja("TIENDAS"), datos = hoja.getDataRange().getValues();
  for (var i=1;i<datos.length;i++) {
    if (String(datos[i][1]).toLowerCase()===String(nombre).toLowerCase()) {
      hoja.getRange(i+1,8).setValue(parseInt(datos[i][7]||0)+1);
      if (!fueGratis && precio && parseFloat(precio)>0) {
        var com  = calcularComision(precio);
        var ant  = parseFloat(datos[i][14]||0);
        var nvo  = Math.max(0,ant-com);
        hoja.getRange(i+1,15).setValue(nvo);
        _registrarMovimiento(nombre,"COMISION",com,"Comisión venta $"+precio,ant,nvo);
      } else if (fueGratis) {
        _registrarMovimiento(nombre,"VENTA_GRATIS",0,"Venta gratuita s/comisión",
          parseFloat(datos[i][14]||0),parseFloat(datos[i][14]||0));
      }
      return "OK";
    }
  }
  return "NO_ENCONTRADO";
}

function guardarCalificacion(tienda, estrellas) {
  obtenerHoja("CALIFICACIONES").appendRow([new Date(),tienda,estrellas]);
  return "OK";
}

function obtenerMetricas(nombre) {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var dT=ss.getSheetByName("TIENDAS")?ss.getSheetByName("TIENDAS").getDataRange().getValues():[];
  var v=0;
  for (var i=1;i<dT.length;i++) {
    if (String(dT[i][1]).toLowerCase()===String(nombre).toLowerCase()) { v=dT[i][7]||0; break; }
  }
  var dC=ss.getSheetByName("CALIFICACIONES")?ss.getSheetByName("CALIFICACIONES").getDataRange().getValues():[];
  var s=0,c=0;
  for (var j=1;j<dC.length;j++) {
    if (String(dC[j][1]).toLowerCase()===String(nombre).toLowerCase()) { s+=parseInt(dC[j][2]); c++; }
  }
  return {ventas:v, promedio:c>0?(s/c).toFixed(1):"Nuevo"};
}

// ══════════════════════════════════════════════════════════
//  EMPLEO
// ══════════════════════════════════════════════════════════
function publicarOfertaEmpleo(obj) {
  // Descontar $5 del saldo de la empresa reclutadora
  if (obj.empresa_login) {
    var hoja = obtenerHoja("EMPRESAS"), datos = hoja.getDataRange().getValues();
    for (var i=1;i<datos.length;i++) {
      if (String(datos[i][1]).toLowerCase()===String(obj.empresa_login).toLowerCase()) {
        var ant = parseFloat(datos[i][11]||0);
        var nvo = Math.max(0,ant-COMISION_VACANTE);
        hoja.getRange(i+1,12).setValue(nvo);
        _registrarMovimientoEmpresa(obj.empresa_login,"COMISION_EMPLEO",COMISION_VACANTE,"Publicación de oferta",ant,nvo);
        // Incrementar contador de publicaciones
        hoja.getRange(i+1,8).setValue(parseInt(datos[i][7]||0)+1);
        break;
      }
    }
  }
  var id="OFR_"+new Date().getTime();
  obtenerHoja("OFERTAS_EMPLEO").appendRow([new Date(),obj.empresa,obj.puesto,obj.sueldo,obj.horario,obj.descripcion,obj.whatsapp,obj.ubicacion,id]);
  return "OK";
}

function publicarBuscoEmpleo(obj) {
  var id="SOL_"+new Date().getTime();
  obtenerHoja("BUSCO_EMPLEO").appendRow([new Date(),obj.nombre,obj.oficio,obj.descripcion,obj.whatsapp,id]);
  return "OK";
}

// ══════════════════════════════════════════════════════════
//  EMPRESAS / RECLUTADORES
// ══════════════════════════════════════════════════════════
function registrarEmpresaServidor(obj) {
  var hoja = obtenerHoja("EMPRESAS"), datos = hoja.getDataRange().getValues();
  for (var i=1;i<datos.length;i++) {
    if (String(datos[i][1]).toLowerCase()===String(obj.nombre).toLowerCase()) return "EXISTE";
  }
  var saldoInicial = regaloActivo() ? SALDO_REGALO : 0;
  var esGratis     = modoGratuitoActivo() ? "SI" : "NO";
  hoja.appendRow([new Date(),obj.nombre,obj.pass,obj.tel,obj.h_abre||"",obj.h_cierra||"",
                  "ACTIVO",0,obj.ubicacion||"",obj.sector||"",obj.whatsapp||obj.tel,
                  saldoInicial,esGratis]);
  if (saldoInicial>0) _registrarMovimientoEmpresa(obj.nombre,"REGALO",saldoInicial,"Saldo de bienvenida",0,saldoInicial);
  return "OK";
}

function loginEmpresa(nombre, pass) {
  var datos = obtenerHoja("EMPRESAS").getDataRange().getValues();
  for (var i=1;i<datos.length;i++) {
    if (String(datos[i][1]).toLowerCase()===String(nombre).toLowerCase() &&
        String(datos[i][2])===String(pass)) {
      var cfg = obtenerConfig();
      return { nombre:String(datos[i][1]), tel:String(datos[i][3]),
               h_abre:String(datos[i][4]||""), h_cierra:String(datos[i][5]||""),
               estatus:String(datos[i][6]||"ACTIVO"), publicaciones:datos[i][7]||0,
               ubicacion:String(datos[i][8]||""), sector:String(datos[i][9]||""),
               whatsapp:String(datos[i][10]||datos[i][3]),
               saldo:parseFloat(datos[i][11]||0),
               es_gratis_origen:String(datos[i][12]||"NO"),
               modo_gratuito_global:String(cfg["modo_gratuito"]||"NO").toUpperCase()==="SI",
               hay_actualizacion:String(cfg["hay_actualizacion"]||"NO").toUpperCase()==="SI",
               url_descarga:String(cfg["url_descarga"]||"") };
    }
  }
  return null;
}

function cambiarPassEmpresa(nombre, actual, nueva, confirm) {
  var hoja = obtenerHoja("EMPRESAS"), datos = hoja.getDataRange().getValues();
  for (var i=1;i<datos.length;i++) {
    if (String(datos[i][1]).toLowerCase()===String(nombre).toLowerCase()) {
      if (String(datos[i][2])!==String(actual)) return {ok:false,msg:"Contraseña actual incorrecta"};
      if (nueva!==confirm) return {ok:false,msg:"Las nuevas contraseñas no coinciden"};
      if (nueva.length<8)  return {ok:false,msg:"Mínimo 8 caracteres"};
      hoja.getRange(i+1,3).setValue(nueva);
      return {ok:true};
    }
  }
  return {ok:false,msg:"Empresa no encontrada"};
}

function actualizarEmpresa(obj) {
  var hoja = obtenerHoja("EMPRESAS"), datos = hoja.getDataRange().getValues();
  for (var i=1;i<datos.length;i++) {
    if (String(datos[i][1]).toLowerCase()===String(obj.nombre).toLowerCase()) {
      if (obj.whatsapp) hoja.getRange(i+1,11).setValue(obj.whatsapp);
      if (obj.h_abre)   hoja.getRange(i+1,5).setValue(obj.h_abre);
      if (obj.h_cierra) hoja.getRange(i+1,6).setValue(obj.h_cierra);
      return "OK";
    }
  }
  return "NO_ENCONTRADO";
}

function obtenerSaldoEmpresa(nombre) {
  var datos = obtenerHoja("EMPRESAS").getDataRange().getValues();
  for (var i=1;i<datos.length;i++) {
    if (String(datos[i][1]).toLowerCase()===String(nombre).toLowerCase()) return parseFloat(datos[i][11]||0);
  }
  return 0;
}

function obtenerMovimientosEmpresa(nombre) {
  var datos = obtenerHoja("MOVIMIENTOS_SALDO").getDataRange().getValues(), lista=[];
  for (var i=1;i<datos.length;i++) {
    if (String(datos[i][1]).toLowerCase()===("EMP:"+String(nombre)).toLowerCase()) {
      lista.push({ fecha:datos[i][0]?Utilities.formatDate(new Date(datos[i][0]),Session.getScriptTimeZone(),"dd/MM/yyyy HH:mm"):"",
                   tipo:String(datos[i][2]||""), monto:parseFloat(datos[i][3]||0),
                   desc:String(datos[i][4]||""), saldo_nuevo:parseFloat(datos[i][6]||0) });
    }
  }
  return lista.reverse();
}

function _registrarMovimientoEmpresa(nombre, tipo, monto, desc, ant, nvo) {
  obtenerHoja("MOVIMIENTOS_SALDO").appendRow([new Date(),"EMP:"+nombre,tipo,monto,desc,ant,nvo]);
}

// Admins pueden ver empresas también
function adminObtenerEmpresas() {
  var datos = obtenerHoja("EMPRESAS").getDataRange().getValues();
  var lista = [];
  for (var i=1;i<datos.length;i++) {
    lista.push({ nombre:String(datos[i][1]||""), estatus:String(datos[i][6]||"ACTIVO"),
                 publicaciones:datos[i][7]||0, saldo:parseFloat(datos[i][11]||0),
                 es_gratis_origen:String(datos[i][12]||"NO") });
  }
  return lista;
}

function adminRecargaSaldoEmpresa(nombre, monto) {
  var hoja = obtenerHoja("EMPRESAS"), datos = hoja.getDataRange().getValues();
  for (var i=1;i<datos.length;i++) {
    if (String(datos[i][1]).toLowerCase()===String(nombre).toLowerCase()) {
      var ant=parseFloat(datos[i][11]||0), nvo=ant+parseFloat(monto);
      hoja.getRange(i+1,12).setValue(nvo);
      _registrarMovimientoEmpresa(nombre,"RECARGA",monto,"Recarga por administrador",ant,nvo);
      return {ok:true,saldo:nvo};
    }
  }
  return {ok:false};
}

function obtenerOfertasEmpleo() {
  var datos=obtenerHoja("OFERTAS_EMPLEO").getDataRange().getValues();
  if (datos.length<2) return [];
  return datos.slice(1).map(function(f){
    return {tipo:"oferta",empresa:String(f[1]||""),puesto:String(f[2]||""),sueldo:String(f[3]||""),
            horario:String(f[4]||""),descripcion:String(f[5]||""),whatsapp:String(f[6]||""),
            ubicacion:String(f[7]||""),id:String(f[8]||"")};
  }).reverse();
}

function obtenerSolicitudesEmpleo() {
  var datos=obtenerHoja("BUSCO_EMPLEO").getDataRange().getValues();
  if (datos.length<2) return [];
  return datos.slice(1).map(function(f){
    return {tipo:"solicitud",nombre:String(f[1]||""),oficio:String(f[2]||""),
            descripcion:String(f[3]||""),whatsapp:String(f[4]||""),id:String(f[5]||"")};
  }).reverse();
}
