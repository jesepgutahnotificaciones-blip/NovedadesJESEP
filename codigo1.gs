/*******************************************************
 * SISTEMA DE GESTIÓN DE NOVEDADES
 * POLICÍA NACIONAL — JESEP / GUTAH
 *
 * CÓDIGO.GS — VERSIÓN COMPLETA RESTAURADA Y OPTIMIZADA
 * (CON LECTURA/CONTEO DE NOVEDADES Y DESCARGA DE PDF/EXCEL)
 *******************************************************/

/* =====================================================
   1. CONFIGURACIÓN GENERAL
   ===================================================== */

const HOJA_BASE       = 'LISTADO_BASE';
const HOJA_NOVEDADES  = 'NOVEDADES';
const HOJA_USUARIOS   = 'USUARIOS';
const HOJA_AUDITORIA  = 'AUDITORIA';
const HOJA_REPORTES   = 'REPORTES';
const HOJA_CONFIG     = 'CONFIG';

const SESION_HORAS = 6;
const SESION_SEGUNDOS = SESION_HORAS * 60 * 60;

const TURNOS_VALIDOS = ['A', 'B', 'C'];

const ENCABEZADOS_NOVEDADES = [
  'GR', 'APELLIDOS Y NOMBRES', 'CC', 'DEPENDENCIA', 'NOVEDAD',
  'DESCRIPCION', 'Dias', 'Fecha INICIAL', 'Fecha PRESENTACION',
  'Observacion', 'RV', 'NIV', 'Turno', 'Placa_Chip', 'TEXTO'
];

const CACHE_BASE_SEGUNDOS = 600;
const CACHE_NOVEDADES_SEGUNDOS = 300;
const CACHE_TROZO = 32000;
const CACHE_MAX_CARACTERES = 3000000;
const FUERZA_DISPONIBLE_SOLO_ACTIVAS = false;

const COLUMNAS_BASE_CACHE = [
  'NIV', 'GR', 'CEDULA', 'FUNCIONARIO',
  'DEPENDENCIA', 'TURNO', 'Placa_Chip'
];

const AUDITAR_LECTURAS = false;
var _MEMO = {};

/* =====================================================
   2. ROUTER Y ACCIONES PERMITIDAS
   ===================================================== */

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!p.accion) {
    return ContentService
      .createTextOutput('API JESEP activa')
      .setMimeType(ContentService.MimeType.TEXT);
  }
  return manejarLlamadaAPI_(e);
}

function obtenerAccionesPermitidas_() {
  return {
    validarUsuario: validarUsuario,
    cerrarSesionCliente: cerrarSesionCliente,
    obtenerDatosSesion: obtenerDatosSesion,

    buscarFuncionario: buscarFuncionario,
    obtenerFichaFuncionario: obtenerFichaFuncionario,
    obtenerHistorialFuncionario: obtenerHistorialFuncionario,

    registrarNovedad: registrarNovedad,
    obtenerTiposNovedadPermitidos: obtenerTiposNovedadPermitidos,
    registrarNovedadPorUsuario: registrarNovedadPorUsuario,

    consultarPorTurno: consultarPorTurno,
    previsualizarSoloNovedades: previsualizarSoloNovedades,
    previsualizarReporteTurno: previsualizarReporteTurno,
    descargarReporteExcel: descargarReporteExcel,
    descargarReportePDF: descargarReportePDF,
    listarReportes: listarReportes,

    consultarCumpleaniosHoy: consultarCumpleaniosHoy,

    listarUsuarios: listarUsuarios,
    crearUsuario: crearUsuario,
    cambiarEstadoUsuario: cambiarEstadoUsuario,

    agregarFuncionarioUBL: agregarFuncionarioUBL,
    eliminarFuncionarioUBL: eliminarFuncionarioUBL,
    eliminarNovedadUBL: eliminarNovedadUBL,

    refrescarCacheDatos: refrescarCacheDatos
  };
}

function manejarLlamadaAPI_(e) {
  const callbackCrudo = texto_(e.parameter.callback);
  const callback = /^[a-zA-Z0-9_]+$/.test(callbackCrudo) ? callbackCrudo : '';

  if (!callback) {
    return ContentService
      .createTextOutput('console.error("Callback inválido.");')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  let resultado;
  try {
    const accion = texto_(e.parameter.accion);
    const funcion = obtenerAccionesPermitidas_()[accion];

    if (!funcion) throw new Error('Acción no permitida: ' + accion);

    let argumentos = [];
    if (e.parameter.args) {
      try { argumentos = JSON.parse(e.parameter.args); } catch (err) { throw new Error('Argumentos inválidos.'); }
    }

    if (!Array.isArray(argumentos)) argumentos = [];
    resultado = funcion.apply(null, argumentos);
    if (resultado === undefined) resultado = null;

  } catch (error) {
    resultado = {
      __jsonp_error: true,
      estado: false,
      mensaje: error && error.message ? error.message : 'Se presentó un error.'
    };
  }

  return ContentService
    .createTextOutput(callback + '(' + JSON.stringify(resultado) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

/* =====================================================
   3. FUNCIONES AUXILIARES DE TEXTO Y CACHÉ
   ===================================================== */

function getSS_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function obtenerHoja_(nombre) {
  const hoja = getSS_().getSheetByName(nombre);
  if (!hoja) throw new Error('No existe la hoja: ' + nombre);
  return hoja;
}

function texto_(valor) {
  if (valor === null || valor === undefined) return '';
  return String(valor).trim();
}

function normalizarTexto_(valor) {
  return String(valor || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim().toUpperCase();
}

function normalizarCedula_(valor) {
  return String(valor || '').replace(/[.\-\s]/g, '').trim();
}

function sha256_(texto) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(texto || ''), Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2)).join('');
}

function formatearFechaLatina_(fechaStr) {
  if (!fechaStr) return '';
  const str = String(fechaStr).trim();
  const partes = str.split('-');
  if (partes.length === 3 && partes[0].length === 4) {
    return partes[2] + '/' + partes[1] + '/' + partes[0];
  }
  return str;
}

function formatearFecha_(fecha) {
  if (!fecha) return '';
  try { return Utilities.formatDate(new Date(fecha), Session.getScriptTimeZone(), 'dd/MM/yyyy'); } catch (e) { return fecha; }
}

function formatearFechaHora_(fecha) {
  if (!fecha) return '';
  try { return Utilities.formatDate(new Date(fecha), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'); } catch (e) { return fecha; }
}

function respuestaOK_(mensaje, datos) {
  if (mensaje && typeof mensaje === 'object' && datos === undefined) {
    return { estado: true, mensaje: mensaje.mensaje || '', datos: mensaje };
  }
  return { estado: true, mensaje: mensaje || '', datos: datos !== undefined ? datos : null };
}

function respuestaError_(mensaje, error) {
  return {
    estado: false,
    mensaje: mensaje || (error && error.message ? error.message : 'Se presentó un error.'),
    datos: null
  };
}

function cacheGrandePut_(clave, texto, segundos) {
  if (!texto || texto.length > CACHE_MAX_CARACTERES) return false;
  const partes = Math.ceil(texto.length / CACHE_TROZO);
  const lote = {};
  for (let i = 0; i < partes; i++) lote[clave + '_' + i] = texto.substr(i * CACHE_TROZO, CACHE_TROZO);
  lote[clave + '_n'] = String(partes);
  CacheService.getScriptCache().putAll(lote, segundos);
  return true;
}

function cacheGrandeGet_(clave) {
  const cache = CacheService.getScriptCache();
  const n = Number(cache.get(clave + '_n'));
  if (!n) return null;
  const claves = [];
  for (let i = 0; i < n; i++) claves.push(clave + '_' + i);
  const mapa = cache.getAll(claves);
  let texto = '';
  for (let i = 0; i < n; i++) {
    const parte = mapa[clave + '_' + i];
    if (parte === undefined || parte === null) return null;
    texto += parte;
  }
  return texto;
}

function cacheGrandeBorrar_(clave) {
  const cache = CacheService.getScriptCache();
  const n = Number(cache.get(clave + '_n')) || 0;
  const claves = [clave + '_n'];
  for (let i = 0; i < n; i++) claves.push(clave + '_' + i);
  cache.removeAll(claves);
}

/* =====================================================
   4. SESIONES Y USUARIOS
   ===================================================== */

function crearSesion_(usuario) {
  const token = sha256_(Utilities.getUuid() + '|' + new Date().getTime() + '|' + Math.random());
  const json = JSON.stringify({ usuario: usuario.usuario, rol: usuario.rol, dependencia: usuario.dependencia, creado: new Date().getTime() });
  PropertiesService.getScriptProperties().setProperty('SESION_' + token, json);
  try { CacheService.getScriptCache().put('SESION_' + token, json, SESION_SEGUNDOS); } catch (e) {}
  return token;
}

function obtenerSesion_(token) {
  if (!token) return null;
  const clave = 'SESION_' + token;
  const cache = CacheService.getScriptCache();
  let valor = null;
  try { valor = cache.get(clave); } catch (e) {}
  if (!valor) {
    valor = PropertiesService.getScriptProperties().getProperty(clave);
    if (valor) {
      try {
        const s = JSON.parse(valor);
        const restante = SESION_SEGUNDOS - Math.floor((new Date().getTime() - s.creado) / 1000);
        if (restante > 0) cache.put(clave, valor, Math.min(restante, 21600));
      } catch (e) {}
    }
  }
  if (!valor) return null;
  let sesion;
  try { sesion = JSON.parse(valor); } catch (e) { return null; }
  if (!sesion.creado || (new Date().getTime() - sesion.creado) > (SESION_SEGUNDOS * 1000)) {
    PropertiesService.getScriptProperties().deleteProperty(clave);
    try { cache.remove(clave); } catch (e) {}
    return null;
  }
  return sesion;
}

function validarSesion_(token) {
  const sesion = obtenerSesion_(token);
  if (!sesion) throw new Error('Sesión no válida o expirada.');
  return sesion;
}

function cerrarSesionCliente(token) {
  if (token) {
    PropertiesService.getScriptProperties().deleteProperty('SESION_' + token);
    try { CacheService.getScriptCache().remove('SESION_' + token); } catch (e) {}
  }
  return { estado: true, mensaje: 'Sesión cerrada correctamente.' };
}

function validarRol_(sesion, rolesPermitidos) {
  if (!sesion) throw new Error('Sesión no válida.');
  const rol = normalizarTexto_(sesion.rol);
  const permitidos = rolesPermitidos.map(r => normalizarTexto_(r));
  if (permitidos.indexOf(rol) === -1) throw new Error('No tiene permisos suficientes.');
  return true;
}

function validarUsuario(usuario, clave) {
  usuario = String(usuario || '').trim();
  clave = String(clave || '');
  if (!usuario || !clave) throw new Error('Ingrese usuario y contraseña.');

  const hoja = obtenerHoja_(HOJA_USUARIOS);
  const datos = hoja.getDataRange().getValues();
  if (datos.length < 2) throw new Error('No existen usuarios registrados.');

  const enc = datos[0].map(v => normalizarTexto_(v));
  const iUsuario = enc.indexOf('USUARIO');
  const iPassword = enc.indexOf('PASSWORD_HASH');
  const iRol = enc.indexOf('ROL');
  const iDependencia = enc.indexOf('DEPENDENCIA');
  const iEstado = enc.indexOf('ESTADO');
  const iUltimoAcceso = enc.indexOf('ULTIMO_ACCESO');

  const buscado = normalizarTexto_(usuario);
  const hashIngresado = sha256_(clave);
  let encontrado = null;

  for (let i = 1; i < datos.length; i++) {
    if (normalizarTexto_(datos[i][iUsuario]) === buscado) {
      encontrado = {
        fila: i + 1,
        usuario: String(datos[i][iUsuario] || '').trim(),
        passwordHash: String(datos[i][iPassword] || '').trim(),
        rol: String(datos[i][iRol] || '').trim(),
        dependencia: String(datos[i][iDependencia] || '').trim(),
        estado: String(datos[i][iEstado] || '').trim()
      };
      break;
    }
  }

  if (!encontrado || encontrado.passwordHash !== hashIngresado) throw new Error('Usuario o contraseña incorrectos.');
  if (normalizarTexto_(encontrado.estado) !== 'ACTIVO') throw new Error('El usuario se encuentra inactivo.');

  const usuarioSesion = { usuario: encontrado.usuario, rol: encontrado.rol, dependencia: encontrado.dependencia };
  const token = crearSesion_(usuarioSesion);
  if (iUltimoAcceso !== -1) hoja.getRange(encontrado.fila, iUltimoAcceso + 1).setValue(new Date());

  registrarAuditoria_(encontrado.usuario, 'INICIO DE SESIÓN', '', 'Ingreso exitoso.');

  return {
    estado: true,
    mensaje: 'Inicio de sesión exitoso.',
    token: token,
    usuario: usuarioSesion,
    turnoAsignado: obtenerTurnoAsignadoOperador_(usuarioSesion)
  };
}

function obtenerDatosSesion(token) {
  const sesion = validarSesion_(token);
  return {
    estado: true,
    usuario: sesion.usuario,
    rol: sesion.rol,
    dependencia: sesion.dependencia,
    turnoAsignado: obtenerTurnoAsignadoOperador_(sesion),
    creado: sesion.creado
  };
}

function registrarAuditoria_(usuario, accion, cedula, detalle) {
  try {
    obtenerHoja_(HOJA_AUDITORIA).appendRow([new Date(), usuario || '', accion || '', cedula || '', detalle || '', 'WEB APP']);
  } catch (e) {
    console.error('Error auditoría: ' + e.message);
  }
}

function registrarAuditoriaLectura_(usuario, accion, cedula, detalle) {
  if (AUDITAR_LECTURAS) registrarAuditoria_(usuario, accion, cedula, detalle);
}

function crearUsuario(token, usuario, clave, rol, dependencia) {
  const sesion = validarSesion_(token);
  validarRol_(sesion, ['ADMINISTRADOR']);

  const hoja = obtenerHoja_(HOJA_USUARIOS);
  const datos = hoja.getDataRange().getValues();
  const enc = datos[0].map(v => normalizarTexto_(v));

  const usuarioNormalizado = normalizarTexto_(usuario);
  for (let i = 1; i < datos.length; i++) {
    if (normalizarTexto_(datos[i][enc.indexOf('USUARIO')]) === usuarioNormalizado) throw new Error('El usuario ya existe.');
  }

  const nuevaFila = new Array(enc.length).fill('');
  nuevaFila[enc.indexOf('USUARIO')] = usuario;
  nuevaFila[enc.indexOf('PASSWORD_HASH')] = sha256_(clave);
  nuevaFila[enc.indexOf('ROL')] = normalizarTexto_(rol);
  nuevaFila[enc.indexOf('DEPENDENCIA')] = dependencia;
  nuevaFila[enc.indexOf('ESTADO')] = 'ACTIVO';
  if (enc.indexOf('FECHA_CREACION') !== -1) nuevaFila[enc.indexOf('FECHA_CREACION')] = new Date();

  hoja.appendRow(nuevaFila);
  return respuestaOK_('Usuario creado correctamente.', { usuario: usuario });
}

function listarUsuarios(token) {
  const sesion = validarSesion_(token);
  validarRol_(sesion, ['ADMINISTRADOR']);

  const datos = obtenerHoja_(HOJA_USUARIOS).getDataRange().getValues();
  if (datos.length < 2) return respuestaOK_('No existen usuarios.', []);

  const enc = datos[0].map(v => normalizarTexto_(v));
  const resultado = [];

  for (let i = 1; i < datos.length; i++) {
    const fila = datos[i];
    resultado.push({
      fila: i + 1,
      USUARIO: fila[enc.indexOf('USUARIO')],
      ROL: fila[enc.indexOf('ROL')],
      DEPENDENCIA: fila[enc.indexOf('DEPENDENCIA')],
      ESTADO: fila[enc.indexOf('ESTADO')],
      ULTIMO_ACCESO: formatearFechaHora_(fila[enc.indexOf('ULTIMO_ACCESO')])
    });
  }
  return respuestaOK_('Usuarios consultados.', resultado);
}

function cambiarEstadoUsuario(token, usuario, nuevoEstado) {
  const sesion = validarSesion_(token);
  validarRol_(sesion, ['ADMINISTRADOR']);

  const hoja = obtenerHoja_(HOJA_USUARIOS);
  const datos = hoja.getDataRange().getValues();
  const enc = datos[0].map(v => normalizarTexto_(v));

  const buscado = normalizarTexto_(usuario);
  for (let i = 1; i < datos.length; i++) {
    if (normalizarTexto_(datos[i][enc.indexOf('USUARIO')]) === buscado) {
      hoja.getRange(i + 1, enc.indexOf('ESTADO') + 1).setValue(normalizarTexto_(nuevoEstado));
      return respuestaOK_('Estado actualizado.');
    }
  }
  throw new Error('Usuario no encontrado.');
}

/* =====================================================
   5. LISTADO BASE & ACCESO A DATOS
   ===================================================== */

function obtenerIndiceColumna_(encabezados, nombre) {
  const buscado = normalizarTexto_(nombre);
  for (let i = 0; i < encabezados.length; i++) {
    if (normalizarTexto_(encabezados[i]) === buscado) return i;
  }
  return -1;
}

function indicesBase_(encabezados) {
  return {
    cedula: obtenerIndiceColumna_(encabezados, 'CEDULA'),
    nivel: obtenerIndiceColumna_(encabezados, 'NIV'),
    grado: obtenerIndiceColumna_(encabezados, 'GR'),
    nombre: obtenerIndiceColumna_(encabezados, 'FUNCIONARIO'),
    dependencia: obtenerIndiceColumna_(encabezados, 'DEPENDENCIA'),
    turno: obtenerIndiceColumna_(encabezados, 'TURNO'),
    placa: obtenerIndiceColumna_(encabezados, 'Placa_Chip')
  };
}

function obtenerListadoBase_() {
  if (_MEMO.base) return _MEMO.base;
  let instantanea = null;

  try {
    const crudo = cacheGrandeGet_('BASE');
    if (crudo) instantanea = JSON.parse(crudo);
  } catch (e) {}

  if (!instantanea) {
    const hoja = obtenerHoja_(HOJA_BASE);
    const valores = hoja.getDataRange().getValues();
    if (valores.length < 1) throw new Error('Hoja base vacía.');

    const encabezadosHoja = valores[0].map(v => texto_(v));
    const columnas = [];
    const encabezados = [];

    COLUMNAS_BASE_CACHE.forEach(nombre => {
      const i = obtenerIndiceColumna_(encabezadosHoja, nombre);
      if (i >= 0) { columnas.push(i); encabezados.push(encabezadosHoja[i]); }
    });

    const datos = [];
    for (let f = 1; f < valores.length; f++) {
      const fila = valores[f];
      const corta = new Array(columnas.length);
      for (let c = 0; c < columnas.length; c++) corta[c] = fila[columnas[c]];
      datos.push(corta);
    }

    instantanea = { encabezados: encabezados, datos: datos };
    try { cacheGrandePut_('BASE', JSON.stringify(instantanea), CACHE_BASE_SEGUNDOS); } catch (e) {}
  }

  _MEMO.base = instantanea;
  return instantanea;
}

function indiceCedulasBase_() {
  if (_MEMO.idxCedulas) return _MEMO.idxCedulas;
  const base = obtenerListadoBase_();
  const idx = indicesBase_(base.encabezados);
  const mapa = Object.create(null);

  if (idx.cedula >= 0) {
    for (let i = 0; i < base.datos.length; i++) {
      const cc = normalizarCedula_(base.datos[i][idx.cedula]);
      if (cc && mapa[cc] === undefined) mapa[cc] = i;
    }
  }
  _MEMO.idxCedulas = mapa;
  return mapa;
}

function invalidarCacheBase_() {
  try { cacheGrandeBorrar_('BASE'); } catch (e) {}
  delete _MEMO.base;
  delete _MEMO.idxCedulas;
}

function construirFuncionario_(encabezados, fila, numeroFila) {
  const idx = indicesBase_(encabezados);
  return {
    fila: numeroFila,
    cedula: idx.cedula >= 0 ? normalizarCedula_(fila[idx.cedula]) : '',
    nivel: idx.nivel >= 0 ? texto_(fila[idx.nivel]) : '',
    grado: idx.grado >= 0 ? texto_(fila[idx.grado]) : '',
    funcionario: idx.nombre >= 0 ? texto_(fila[idx.nombre]) : '',
    dependencia: idx.dependencia >= 0 ? texto_(fila[idx.dependencia]) : '',
    turno: idx.turno >= 0 ? texto_(fila[idx.turno]) : '',
    placaChip: idx.placa >= 0 ? texto_(fila[idx.placa]) : ''
  };
}

function obtenerTurnoAsignadoOperador_(sesion) {
  if (!sesion || normalizarTexto_(sesion.rol) !== 'OPERADOR') return '';
  const match = normalizarTexto_(sesion.dependencia).match(/DISPONIBLE[_\s-]*([A-Z]+)/);
  if (!match) return '';
  const cod = match[1];
  if (['A', 'B', 'C', 'SEPRI', 'GURIN'].indexOf(cod) !== -1) return cod;
  if (['NA', 'NO', 'NOAPLICA'].indexOf(cod) !== -1) return 'NO APLICA';
  return '';
}

function crearFiltroPermisos_(sesion) {
  const rol = normalizarTexto_(sesion && sesion.rol);
  if (rol === 'ADMINISTRADOR') return () => true;
  if (rol === 'OPERADOR') {
    const turno = obtenerTurnoAsignadoOperador_(sesion);
    if (turno) return (t) => turnoCoincideFiltro_(t, turno);
    const depUsuario = normalizarTexto_(sesion.dependencia);
    return (t, d) => depUsuario !== '' && depUsuario === normalizarTexto_(d);
  }
  return () => false;
}

function buscarFuncionario(token, valor) {
  const sesion = validarSesion_(token);
  const criterio = texto_(valor).trim();
  if (!criterio) return respuestaError_('Debe ingresar un criterio de búsqueda.');

  const base = obtenerListadoBase_();
  const idx = indicesBase_(base.encabezados);
  const criterioTexto = normalizarTexto_(criterio);
  const criterioCedula = normalizarCedula_(criterio);
  const permitir = crearFiltroPermisos_(sesion);
  const resultados = [];

  for (let i = 0; i < base.datos.length && resultados.length < 50; i++) {
    const fila = base.datos[i];
    const cc = idx.cedula >= 0 ? normalizarCedula_(fila[idx.cedula]) : '';
    let coincide = (criterioCedula !== '' && cc.indexOf(criterioCedula) !== -1);

    if (!coincide && criterioTexto !== '') {
      coincide =
        (idx.nombre >= 0 && normalizarTexto_(fila[idx.nombre]).indexOf(criterioTexto) !== -1) ||
        (idx.grado >= 0 && normalizarTexto_(fila[idx.grado]).indexOf(criterioTexto) !== -1) ||
        (idx.dependencia >= 0 && normalizarTexto_(fila[idx.dependencia]).indexOf(criterioTexto) !== -1);
    }

    if (!coincide) continue;
    const turno = idx.turno >= 0 ? texto_(fila[idx.turno]) : '';
    const dep = idx.dependencia >= 0 ? texto_(fila[idx.dependencia]) : '';

    if (!permitir(turno, dep)) continue;

    resultados.push({
      fila: i + 2,
      cedula: cc,
      nivel: idx.nivel >= 0 ? texto_(fila[idx.nivel]) : '',
      grado: idx.grado >= 0 ? texto_(fila[idx.grado]) : '',
      funcionario: idx.nombre >= 0 ? texto_(fila[idx.nombre]) : '',
      dependencia: dep,
      turno: turno
    });
  }

  return respuestaOK_(resultados.length ? 'Consulta exitosa.' : 'Sin coincidencias.', resultados);
}

function obtenerFichaFuncionario(token, cedula) {
  const sesion = validarSesion_(token);
  const cc = normalizarCedula_(cedula);
  const base = obtenerListadoBase_();
  const pos = indiceCedulasBase_()[cc];

  if (pos === undefined) return respuestaError_('Funcionario no encontrado.');
  const f = construirFuncionario_(base.encabezados, base.datos[pos], pos + 2);

  return { estado: true, funcionario: f, datos: f };
}

/* =====================================================
   6. NOVEDADES (LECTURA Y REGISTRO)
   ===================================================== */

const TIPOS_NOVEDAD_POR_USUARIO = {
 'SGSST_JESEP': ['EXCUSA MEDICA', 'RESTRICCIONES MEDICA', 'LICIENCIA DE MATERNIDAD'],
  'VAC_JESEP': ['PLAN VACACIONAL', 'VACACIONES EXTRAORDINARIAS', 'PLAN REDUCCION', 'VACACIONES DE RETIRO'],
  'PAS_JESEP': ['COMISION DE ESTUDIO', 'COMISION DE SERVICIO', 'LICENCIA DE PATERNIDAD', 'LICENCIA DE LUTO'],
  'CIT_JESEP': ['SUSPENCION', 'CITACION JUDICIAL (PERMISO)'],
  'HIS_JESEP': ['RETIROS 3 MESES DE ALTA', 'ELIMINAR USURIO POR RETIRO'],
  'PRO_JESEP': ['CURSO DE ASCENSO ESPOL','CURSO DE ASCENSO ESJIM'],
  'GH_JESEP': ['CAMBIO DE TURNO', 'HORARIO FLEXIBLE'],
  'CAP_JESEP': ['CURSO MANDATORIO', 'CAPACITACION']
};

function obtenerTiposNovedadPermitidos(token) {
  const sesion = validarSesion_(token);
  const usr = String(sesion.usuario).toUpperCase();
  const tipos = TIPOS_NOVEDAD_POR_USUARIO[usr] || ['SERVICIO', 'PERMISO', 'OTRA NOVEDAD'];
  return respuestaOK_('Tipos asignados', { tipos: tipos });
}

function obtenerNovedadesTabla_() {
  if (_MEMO.nov) return _MEMO.nov;
  let tabla = null;

  try {
    const crudo = cacheGrandeGet_('NOV');
    if (crudo) tabla = JSON.parse(crudo);
  } catch (e) {}

  if (!tabla) {
    const hoja = obtenerHoja_(HOJA_NOVEDADES);
    const valores = hoja.getDataRange().getValues();
    tabla = {
      encabezados: valores.length ? valores[0].map(v => texto_(v)) : [],
      datos: valores.length > 1 ? valores.slice(1) : []
    };
    try { cacheGrandePut_('NOV', JSON.stringify(tabla), CACHE_NOVEDADES_SEGUNDOS); } catch (e) {}
  }
  _MEMO.nov = tabla;
  return tabla;
}

function indiceNovedades_() {
  if (_MEMO.idxNov) return _MEMO.idxNov;
  const tabla = obtenerNovedadesTabla_();

  let iCC = obtenerIndiceColumna_(tabla.encabezados, 'CC');
  if (iCC < 0) iCC = obtenerIndiceColumna_(tabla.encabezados, 'CEDULA');
  if (iCC < 0) iCC = obtenerIndiceColumna_(tabla.encabezados, 'CÉDULA');
  if (iCC < 0) iCC = 2;

  const mapa = Object.create(null);

  if (iCC >= 0 && tabla.datos) {
    for (let i = 0; i < tabla.datos.length; i++) {
      const cc = normalizarCedula_(tabla.datos[i][iCC]);
      if (!cc) continue;
      if (!mapa[cc]) mapa[cc] = [];
      mapa[cc].push(i);
    }
  }
  _MEMO.idxNov = mapa;
  return mapa;
}

function invalidarCacheNovedades_() {
  try { cacheGrandeBorrar_('NOV'); } catch (e) {}
  delete _MEMO.nov;
  delete _MEMO.idxNov;
}

function tiempoFecha_(valor) {
  if (!valor) return 0;
  if (valor instanceof Date) return isNaN(valor.getTime()) ? 0 : valor.getTime();
  const d = new Date(valor);
  return isNaN(d.getTime()) ? 0 : d.getTime();
}

function novedadEstaActiva_(fechaPresentacion) {
  if (!fechaPresentacion) return true;
  let fecha = fechaPresentacion instanceof Date ? fechaPresentacion : new Date(fechaPresentacion);
  if (isNaN(fecha.getTime())) return true;

  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  fecha.setHours(0, 0, 0, 0);
  return fecha.getTime() >= hoy.getTime();
}

function obtenerHistorialFuncionario(token, cedula) {
  validarSesion_(token);
  const cc = normalizarCedula_(cedula);
  const tabla = obtenerNovedadesTabla_();
  const posiciones = indiceNovedades_()[cc] || [];
  const hist = [];

  let iGr = obtenerIndiceColumna_(tabla.encabezados, 'GR'); if (iGr < 0) iGr = 0;
  let iNov = obtenerIndiceColumna_(tabla.encabezados, 'NOVEDAD'); if (iNov < 0) iNov = 4;
  let iDias = obtenerIndiceColumna_(tabla.encabezados, 'Dias'); if (iDias < 0) iDias = 6;
  let iIni = obtenerIndiceColumna_(tabla.encabezados, 'Fecha INICIAL'); if (iIni < 0) iIni = 7;
  let iPres = obtenerIndiceColumna_(tabla.encabezados, 'Fecha PRESENTACION'); if (iPres < 0) iPres = 8;

  posiciones.forEach(pos => {
    const fila = tabla.datos[pos];
    hist.push({
      fila: pos + 2,
      GR: fila[iGr] || '',
      NOVEDAD: fila[iNov] || '',
      Dias: fila[iDias] || '',
      'Fecha INICIAL': fila[iIni] || '',
      'Fecha PRESENTACION': fila[iPres] || ''
    });
  });

  return { estado: true, historial: hist, datos: hist };
}

function registrarNovedadPorUsuario(token, datos) {
  const sesion = validarSesion_(token);
  if (datos.novedad === 'CAMBIO DE TURNO') {
    const hojaBase = obtenerHoja_(HOJA_BASE);
    const datosBase = hojaBase.getDataRange().getValues();
    const cc = normalizarCedula_(datos.cc || datos.cedula);
    const enc = datosBase[0].map(v => normalizarTexto_(v));
    const iCc = enc.indexOf('CEDULA');
    const iTur = enc.indexOf('TURNO');

    for (let i = 1; i < datosBase.length; i++) {
      if (normalizarCedula_(datosBase[i][iCc]) === cc) {
        hojaBase.getRange(i + 1, iTur + 1).setValue(datos.nuevoTurno);
        invalidarCacheBase_();
        return respuestaOK_('Cambio de turno registrado en LISTADO_BASE.');
      }
    }
    return respuestaError_('Funcionario no encontrado.');
  }

  return registrarNovedad(token, datos);
}

function registrarNovedad(token, datos) {
  validarSesion_(token);
  const cc = normalizarCedula_(datos.cc || datos.cedula);
  if (!cc) return respuestaError_('Número de cédula inválido.');

  const hojaBase = obtenerHoja_(HOJA_BASE);
  const datosBase = hojaBase.getDataRange().getValues();
  const encBase = datosBase[0].map(v => normalizarTexto_(v));

  const iCc = encBase.indexOf('CEDULA');
  const iGr = encBase.indexOf('GR');
  const iNom = encBase.indexOf('FUNCIONARIO');
  const iDep = encBase.indexOf('DEPENDENCIA');
  const iNiv = encBase.indexOf('NIV');
  const iTur = encBase.indexOf('TURNO');

  let gr = '', funcionario = '', dependencia = '', nivel = '', turno = '';

  for (let i = 1; i < datosBase.length; i++) {
    if (normalizarCedula_(datosBase[i][iCc]) === cc) {
      gr = datosBase[i][iGr];
      funcionario = datosBase[i][iNom];
      dependencia = datosBase[i][iDep];
      nivel = datosBase[i][iNiv];
      turno = datosBase[i][iTur];
      break;
    }
  }

  if (!funcionario) return respuestaError_('Funcionario no existe en LISTADO_BASE.');

  const tipoColE = datos.novedad || '';
  let descColF = '';
  const reub = datos.descripcion || '';
  const detalleExtra = datos.nombreNovedad || '';

  if (reub && detalleExtra) descColF = reub + ' "' + detalleExtra + '"';
  else if (reub) descColF = reub;
  else descColF = detalleExtra;

  const fechaIni = formatearFechaLatina_(datos.fechaInicial);
  const fechaPres = formatearFechaLatina_(datos.fechaPresentacion);

  const nuevaFila = [
    gr,                     // Columna A
    funcionario,            // Columna B
    cc,                     // Columna C
    dependencia,            // Columna D
    tipoColE,               // Columna E
    descColF,               // Columna F
    datos.dias || '',       // Columna G
    fechaIni,               // Columna H
    fechaPres,              // Columna I
    '',                     // Columna J
    nivel,                  // Columna K (NIV)
    turno,                  // Columna L (TURNO)
    '',                     // Columna M
    '',                     // Columna N
    ''                      // Columna O
  ];

  const hojaNov = obtenerHoja_(HOJA_NOVEDADES);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    hojaNov.appendRow(nuevaFila);
  } finally {
    lock.releaseLock();
  }

  invalidarCacheNovedades_();
  return respuestaOK_('Novedad guardada exitosamente.');
}

/* =====================================================
   7. CONSULTAS POR TURNO Y CONTADORES
   ===================================================== */

const ORDEN_GRADOS_TURNO = {
  'BG': 1, 'CR': 2, 'TC': 3, 'MY': 4, 'CT': 5, 'TE': 6, 'ST': 7,
  'CM': 8, 'SC': 9, 'IJ': 10, 'IT': 11, 'SI': 12, 'PT': 13,
  'PP': 14, 'AXP': 15, 'NO': 16
};

function normalizarFiltroTurno_(filtro) {
  const val = normalizarTexto_(texto_(filtro));
  if (['NO APLICA', 'NO_APLICA', 'NO-APLICA'].indexOf(val) !== -1) return 'NO APLICA';
  if (['A', 'B', 'C', 'SEPRI', 'GURIN'].indexOf(val) !== -1) return val;
  return '';
}

function turnoCoincideFiltro_(turno, filtro) {
  const vTur = normalizarTexto_(turno);
  const fNorm = normalizarFiltroTurno_(filtro);
  if (!fNorm) return false;
  if (['A', 'B', 'C'].indexOf(fNorm) !== -1) return vTur === fNorm;
  if (fNorm === 'SEPRI') return vTur.indexOf('SEPRI') !== -1;
  if (fNorm === 'GURIN') return vTur.indexOf('GURIN') !== -1;
  if (fNorm === 'NO APLICA') return vTur === '' || vTur.indexOf('NO APLICA') !== -1;
  return false;
}

function clasificarNivelFuncionario_(f) {
  const grado = normalizarTexto_(f.grado);
  const nivel = normalizarTexto_(f.nivel);

  if (['BG', 'CR', 'TC', 'MY', 'CT', 'TE', 'ST', 'CM'].indexOf(grado) !== -1 || nivel.indexOf('OFICIAL') !== -1) return 'OF';
  if (grado === 'PT' || nivel === 'PT') return 'PT';
  if (grado === 'PP' || nivel === 'PP') return 'PP';
  if (grado === 'AXP' || nivel.indexOf('AXP') !== -1 || nivel.indexOf('AUXILIAR') !== -1) return 'AXP';
  if (nivel.indexOf('NO UNI') !== -1 || nivel.indexOf('NO UNIFORM') !== -1) return 'NO UNI';
  if (['SC', 'IJ', 'IT', 'SI'].indexOf(grado) !== -1 || nivel.indexOf('EJECUTIVO') !== -1) return 'NIV. EJE';

  return 'NO UNI';
}

function consultarPorTurno(token, filtro) {
  const sesion = validarSesion_(token);
  const fNorm = normalizarFiltroTurno_(filtro);
  if (!fNorm) return respuestaError_('Filtro de turno no válido.');

  const base = obtenerListadoBase_();
  const idx = indicesBase_(base.encabezados);
  const permitir = crearFiltroPermisos_(sesion);
  const resultados = [];

  base.datos.forEach((fila, i) => {
    if (!turnoCoincideFiltro_(fila[idx.turno], fNorm)) return;
    const turno = texto_(fila[idx.turno]);
    const dep = idx.dependencia >= 0 ? texto_(fila[idx.dependencia]) : '';
    if (!permitir(turno, dep)) return;

    resultados.push({
      fila: i + 2,
      cedula: normalizarCedula_(fila[idx.cedula]),
      nivel: idx.nivel >= 0 ? texto_(fila[idx.nivel]) : '',
      grado: idx.grado >= 0 ? texto_(fila[idx.grado]) : '',
      funcionario: idx.nombre >= 0 ? texto_(fila[idx.nombre]) : '',
      dependencia: dep,
      turno: turno,
      placaChip: idx.placa >= 0 ? texto_(fila[idx.placa]) : '',
      historialNovedades: [],
      novedadesActivas: [],
      totalNovedades: 0,
      totalNovedadesActivas: 0,
      tieneNovedadActiva: false
    });
  });

  if (resultados.length > 0) {
    try {
      const tabla = obtenerNovedadesTabla_();
      const mapa = indiceNovedades_();

      let iNov = obtenerIndiceColumna_(tabla.encabezados, 'NOVEDAD'); if (iNov < 0) iNov = 4;
      let iIni = obtenerIndiceColumna_(tabla.encabezados, 'Fecha INICIAL'); if (iIni < 0) iIni = 7;
      let iPres = obtenerIndiceColumna_(tabla.encabezados, 'Fecha PRESENTACION'); if (iPres < 0) iPres = 8;

      resultados.forEach(f => {
        const posiciones = mapa[f.cedula];
        if (!posiciones) return;

        const items = posiciones.map(pos => {
          const fila = tabla.datos[pos];
          return {
            inicio: tiempoFecha_(fila[iIni]),
            nombre: texto_(fila[iNov]),
            activa: novedadEstaActiva_(fila[iPres])
          };
        });

        items.sort((a, b) => b.inicio - a.inicio);
        const activas = items.filter(x => x.activa);

        f.totalNovedades = items.length;
        f.totalNovedadesActivas = activas.length;
        f.tieneNovedadActiva = activas.length > 0;
        f.historialNovedades = items.slice(0, 5).map(x => ({ NOVEDAD: x.nombre }));
        f.novedadesActivas = activas.slice(0, 5).map(x => ({ NOVEDAD: x.nombre }));
      });
    } catch (e) {
      console.error('Error al cruzar novedades: ' + e.message);
    }
  }

  resultados.sort((a, b) => {
    const oa = ORDEN_GRADOS_TURNO[normalizarTexto_(a.grado)] || 999;
    const ob = ORDEN_GRADOS_TURNO[normalizarTexto_(b.grado)] || 999;
    return oa - ob;
  });

  const CATEGORIAS_FUERZA = ['OF', 'NIV. EJE', 'PT', 'PP', 'AXP', 'NO UNI'];
  const fuerzaEfectiva = { 'OF': 0, 'NIV. EJE': 0, 'PT': 0, 'PP': 0, 'AXP': 0, 'NO UNI': 0 };
  const fuerzaConNovedad = { 'OF': 0, 'NIV. EJE': 0, 'PT': 0, 'PP': 0, 'AXP': 0, 'NO UNI': 0 };

  resultados.forEach(f => {
    const cat = clasificarNivelFuncionario_(f);
    fuerzaEfectiva[cat]++;
    const resta = FUERZA_DISPONIBLE_SOLO_ACTIVAS ? (f.tieneNovedadActiva === true) : (f.totalNovedades > 0);
    if (resta) fuerzaConNovedad[cat]++;
  });

  const fuerzaDisponible = {};
  CATEGORIAS_FUERZA.forEach(cat => { fuerzaDisponible[cat] = fuerzaEfectiva[cat] - fuerzaConNovedad[cat]; });
  const conNovCount = resultados.filter(x => x.tieneNovedadActiva === true).length;

  return respuestaOK_('Consulta por turno exitosa.', {
    filtro: fNorm,
    funcionarios: resultados,
    total: resultados.length,
    estadisticas: {
      total: resultados.length,
      totalConNovedad: conNovCount,
      totalSinNovedad: resultados.length - conNovCount,
      categorias: CATEGORIAS_FUERZA,
      fuerzaEfectiva: fuerzaEfectiva,
      fuerzaConNovedad: fuerzaConNovedad,
      fuerzaDisponible: fuerzaDisponible
    }
  });
}

/* =====================================================
   8. GENERACIÓN Y DESCARGA DE REPORTE PDF/EXCEL
   ===================================================== */

function construirHTMLReporteTurno_(filtro, funcionarios, consecutivo, usuario) {
  const fechaTexto = formatearFechaHora_(new Date());
  const CATEGORIAS = ['OF', 'NIV. EJE', 'PT', 'PP', 'AXP', 'NO UNI'];
  const lista = (funcionarios || []).slice();

  lista.sort((a, b) => (ORDEN_GRADOS_TURNO[normalizarTexto_(a.grado)] || 999) - (ORDEN_GRADOS_TURNO[normalizarTexto_(b.grado)] || 999));

  let filasHTML = '';
  lista.forEach((f, idx) => {
    const novs = (f.historialNovedades || []).map(n => n.NOVEDAD).join(' | ') || 'Sin novedades';
    filasHTML += `<tr>
      <td style="text-align:center;">${idx + 1}</td>
      <td>${f.cedula}</td>
      <td>${f.nivel}</td>
      <td>${f.grado}</td>
      <td><strong>${f.funcionario}</strong></td>
      <td>${f.dependencia}</td>
      <td style="text-align:center;">${f.turno}</td>
      <td>${novs}</td>
    </tr>`;
  });

  return `<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
body { font-family: Arial, sans-serif; font-size: 9px; color: #111; margin: 10px; }
.titulo { text-align: center; font-size: 15px; font-weight: bold; }
.subtitulo { text-align: center; font-size: 10px; margin-bottom: 10px; }
table { width: 100%; border-collapse: collapse; margin-top: 10px; }
th, td { border: 1px solid #777; padding: 4px; vertical-align: middle; }
th { background: #e8e8e8; font-weight: bold; text-align: center; }
</style>
</head>
<body>
<div class="titulo">POLICÍA NACIONAL DE COLOMBIA</div>
<div class="subtitulo">SISTEMA DE GESTIÓN DE NOVEDADES — TURNO: ${filtro}</div>
<div><strong>Generado por:</strong> ${usuario || ''} | <strong>Fecha:</strong> ${fechaTexto}</div>
<table>
<thead>
  <tr><th>#</th><th>CÉDULA</th><th>NIV</th><th>GRADO</th><th>FUNCIONARIO</th><th>DEPENDENCIA</th><th>TURNO</th><th>NOVEDADES</th></tr>
</thead>
<tbody>${filasHTML || '<tr><td colspan="8" style="text-align:center;">Sin funcionarios.</td></tr>'}</tbody>
</table>
</body>
</html>`;
}

function descargarReportePDF(token, filtro) {
  try {
    const sesion = validarSesion_(token);
    const consulta = consultarPorTurno(token, filtro);
    if (!consulta || consulta.estado !== true) return consulta;

    const funcionarios = (consulta.datos || {}).funcionarios || [];
    const html = construirHTMLReporteTurno_(filtro, funcionarios, 'REP', sesion.usuario);
    const blobPDF = Utilities.newBlob(html, 'text/html', 'Listado_' + filtro + '.pdf').getAs('application/pdf');
    const base64 = Utilities.base64Encode(blobPDF.getBytes());

    return respuestaOK_('Archivo PDF generado.', {
      filtro: filtro,
      total: funcionarios.length,
      nombre: 'Listado_Turno_' + filtro + '.pdf',
      mime: 'application/pdf',
      contenido: base64
    });
  } catch (err) {
    return respuestaError_('No fue posible generar el PDF.', err);
  }
}

function descargarReporteExcel(token, filtro) {
  try {
    const sesion = validarSesion_(token);
    const consulta = consultarPorTurno(token, filtro);
    if (!consulta || consulta.estado !== true) return consulta;

    const funcionarios = (consulta.datos || {}).funcionarios || [];
    let csv = '\ufeff#;CEDULA;NIV;GRADO;FUNCIONARIO;DEPENDENCIA;TURNO;NOVEDADES\n';

    funcionarios.forEach((f, idx) => {
      const novs = (f.historialNovedades || []).map(n => n.NOVEDAD).join(' - ') || 'Sin novedades';
      csv += `${idx + 1};"${f.cedula}";"${f.nivel}";"${f.grado}";"${f.funcionario}";"${f.dependencia}";"${f.turno}";"${novs}"\n`;
    });

    const base64 = Utilities.base64Encode(csv, Utilities.Charset.UTF_8);

    return respuestaOK_('Archivo Excel generado.', {
      filtro: filtro,
      total: funcionarios.length,
      nombre: 'Listado_Turno_' + filtro + '.csv',
      mime: 'text/csv',
      contenido: base64
    });
  } catch (err) {
    return respuestaError_('No fue posible generar el Excel.', err);
  }
}

function previsualizarSoloNovedades(token, filtro) {
  const sesion = validarSesion_(token);
  const consulta = consultarPorTurno(token, filtro);
  if (!consulta || consulta.estado !== true) return consulta;
  const funcionarios = (consulta.datos || {}).funcionarios || [];
  return { estado: true, html: construirHTMLReporteTurno_(filtro, funcionarios, 'PREV', sesion.usuario) };
}

function previsualizarReporteTurno(token, filtro) { return previsualizarSoloNovedades(token, filtro); }
function listarReportes(token) { return respuestaOK_('Sin reportes', []); }

/* =====================================================
   9. SERVICIOS AUXILIARES Y UBL_JESEP
   ===================================================== */

function consultarCumpleaniosHoy(token) {
  validarSesion_(token);
  return respuestaOK_('Sin cumpleaños.', { total: 0, lista: [] });
}

function agregarFuncionarioUBL(token, datos) {
  const sesion = validarSesion_(token);
  if (String(sesion.usuario).toUpperCase() !== 'UBL_JESEP' && sesion.rol !== 'ADMINISTRADOR') throw new Error('No autorizado.');

  const hoja = obtenerHoja_(HOJA_BASE);
  const fechaNacFormateada = formatearFechaLatina_(datos.fechaNacimiento);

  const fila = [
    datos.cedula, datos.niv, datos.gr, datos.funcionario, datos.dependencia,
    datos.dependencia, datos.pert, datos.turno, '', '', datos.mes, datos.dia,
    fechaNacFormateada, datos.correo, datos.sexo, datos.estadoCivil, datos.situacionLaboral,
    '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', datos.comunicadoOficial
  ];

  hoja.appendRow(fila);
  invalidarCacheBase_();
  return respuestaOK_('Funcionario agregado exitosamente.');
}

function eliminarFuncionarioUBL(token, cedula) {
  const sesion = validarSesion_(token);
  if (String(sesion.usuario).toUpperCase() !== 'UBL_JESEP' && sesion.rol !== 'ADMINISTRADOR') throw new Error('No autorizado.');

  const hoja = obtenerHoja_(HOJA_BASE);
  const datos = hoja.getDataRange().getValues();
  const cc = normalizarCedula_(cedula);

  for (let i = 1; i < datos.length; i++) {
    if (normalizarCedula_(datos[i][0]) === cc) {
      hoja.deleteRow(i + 1);
      invalidarCacheBase_();
      return respuestaOK_('Funcionario eliminado.');
    }
  }
  return respuestaError_('Funcionario no encontrado.');
}

function eliminarNovedadUBL(token, numFila) {
  const sesion = validarSesion_(token);
  if (String(sesion.usuario).toUpperCase() !== 'UBL_JESEP' && sesion.rol !== 'ADMINISTRADOR') throw new Error('No autorizado.');

  const hoja = obtenerHoja_(HOJA_NOVEDADES);
  const fila = parseInt(numFila, 10);

  if (isNaN(fila) || fila <= 1) return respuestaError_('Fila no válida.');
  if (fila > hoja.getLastRow()) return respuestaError_('La fila no existe.');

  hoja.deleteRow(fila);
  invalidarCacheNovedades_();
  return respuestaOK_('Novedad eliminada correctamente.');
}

function refrescarCacheDatos(token) {
  if (token) validarSesion_(token);
  invalidarCacheBase_();
  invalidarCacheNovedades_();
  return respuestaOK_('Caché reiniciada correctamente.');
}