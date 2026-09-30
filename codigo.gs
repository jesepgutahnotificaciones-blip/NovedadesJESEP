/*******************************************************
 * SISTEMA DE GESTIÓN DE NOVEDADES
 * POLICÍA NACIONAL — JESEP / GUTAH
 * CÓDIGO.GS — VERSIÓN AJUSTADA (COLUMNA K Y FORMATO FECHA DD/MM/YYYY)
 *******************************************************/

const HOJA_BASE       = 'LISTADO_BASE';
const HOJA_NOVEDADES  = 'NOVEDADES';
const HOJA_USUARIOS   = 'USUARIOS';
const HOJA_AUDITORIA  = 'AUDITORIA';
const HOJA_REPORTES   = 'REPORTES';

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!p.accion) {
    return ContentService.createTextOutput('API JESEP activa').setMimeType(ContentService.MimeType.TEXT);
  }
  return manejarLlamadaAPI_(e);
}

function obtenerAccionesPermitidas_() {
  return {
    validarUsuario: validarUsuario,
    cerrarSesionCliente: cerrarSesionCliente,
    buscarFuncionario: buscarFuncionario,
    obtenerFichaFuncionario: obtenerFichaFuncionario,
    obtenerHistorialFuncionario: obtenerHistorialFuncionario,
    registrarNovedad: registrarNovedad,
    obtenerTiposNovedadPermitidos: obtenerTiposNovedadPermitidos,
    registrarNovedadPorUsuario: registrarNovedadPorUsuario,
    previsualizarSoloNovedades: previsualizarSoloNovedades,
    agregarFuncionarioUBL: agregarFuncionarioUBL,
    eliminarFuncionarioUBL: eliminarFuncionarioUBL
  };
}

function manejarLlamadaAPI_(e) {
  const callbackCrudo = texto_(e.parameter.callback);
  const callback = /^[a-zA-Z0-9_]+$/.test(callbackCrudo) ? callbackCrudo : '';

  if (!callback) {
    return ContentService.createTextOutput('console.error("Callback inválido.");').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  let resultado;
  try {
    const accion = texto_(e.parameter.accion);
    const funcion = obtenerAccionesPermitidas_()[accion];
    if (!funcion) throw new Error('Acción no permitida: ' + accion);

    let argumentos = [];
    if (e.parameter.args) argumentos = JSON.parse(e.parameter.args);
    if (!Array.isArray(argumentos)) argumentos = [];

    resultado = funcion.apply(null, argumentos);
    if (resultado === undefined) resultado = null;
  } catch (error) {
    resultado = { __jsonp_error: true, estado: false, mensaje: error.message };
  }

  return ContentService.createTextOutput(callback + '(' + JSON.stringify(resultado) + ');').setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function getSS_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function obtenerHoja_(nombre) { return getSS_().getSheetByName(nombre); }

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

// Función auxiliar para formatear cadenas YYYY-MM-DD a DD/MM/YYYY
function formatearFechaLatina_(fechaStr) {
  if (!fechaStr) return '';
  const str = String(fechaStr).trim();
  const partes = str.split('-');
  if (partes.length === 3 && partes[0].length === 4) {
    return partes[2] + '/' + partes[1] + '/' + partes[0];
  }
  return str;
}

function sha256_(texto) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(texto || ''), Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2)).join('');
}

function respuestaOK_(mensaje, datos) {
  return { estado: true, mensaje: mensaje || '', datos: datos !== undefined ? datos : null };
}

function respuestaError_(mensaje) {
  return { estado: false, mensaje: mensaje || 'Error en la operación.' };
}

function crearSesion_(usuario) {
  const token = sha256_(Utilities.getUuid() + '|' + new Date().getTime());
  const json = JSON.stringify({ usuario: usuario.usuario, rol: usuario.rol, dependencia: usuario.dependencia, creado: new Date().getTime() });
  PropertiesService.getScriptProperties().setProperty('SESION_' + token, json);
  return token;
}

function obtenerSesion_(token) {
  if (!token) return null;
  const valor = PropertiesService.getScriptProperties().getProperty('SESION_' + token);
  if (!valor) return null;
  return JSON.parse(valor);
}

function validarSesion_(token) {
  const sesion = obtenerSesion_(token);
  if (!sesion) throw new Error('Sesión vencida o no válida.');
  return sesion;
}

function cerrarSesionCliente(token) {
  if (token) PropertiesService.getScriptProperties().deleteProperty('SESION_' + token);
  return { estado: true };
}

function validarUsuario(usuario, clave) {
  const hoja = obtenerHoja_(HOJA_USUARIOS);
  const datos = hoja.getDataRange().getValues();
  const bus = normalizarTexto_(usuario);
  const pass = sha256_(clave);

  for (let i = 1; i < datos.length; i++) {
    if (normalizarTexto_(datos[i][0]) === bus && datos[i][1] === pass) {
      const u = { usuario: datos[i][0], rol: datos[i][2], dependencia: datos[i][3] };
      const token = crearSesion_(u);
      return { estado: true, token: token, usuario: u };
    }
  }
  throw new Error('Usuario o contraseña incorrectos.');
}

function buscarFuncionario(token, valor) {
  validarSesion_(token);
  const hoja = obtenerHoja_(HOJA_BASE);
  const datos = hoja.getDataRange().getValues();
  if (datos.length < 2) return respuestaOK_('Resultados de búsqueda', []);

  const enc = datos[0].map(v => normalizarTexto_(v));
  const iCc = enc.indexOf('CEDULA') >= 0 ? enc.indexOf('CEDULA') : 0;
  const iGr = enc.indexOf('GR') >= 0 ? enc.indexOf('GR') : 2;
  const iNom = enc.indexOf('FUNCIONARIO') >= 0 ? enc.indexOf('FUNCIONARIO') : 3;
  const iDep = enc.indexOf('DEPENDENCIA') >= 0 ? enc.indexOf('DEPENDENCIA') : (enc.indexOf('DEP') >= 0 ? enc.indexOf('DEP') : 4);
  const iPert = enc.indexOf('PERT') >= 0 ? enc.indexOf('PERT') : 6;
  const iTur = enc.indexOf('TURNO') >= 0 ? enc.indexOf('TURNO') : 7;

  const crit = normalizarTexto_(valor);
  const res = [];

  for (let i = 1; i < datos.length; i++) {
    const cc = normalizarCedula_(datos[i][iCc]);
    const nom = normalizarTexto_(datos[i][iNom]);
    const dep = normalizarTexto_(datos[i][iDep]);
    const pert = normalizarTexto_(datos[i][iPert]);
    const gr = normalizarTexto_(datos[i][iGr]);

    if (cc.indexOf(crit) !== -1 || nom.indexOf(crit) !== -1 || dep.indexOf(crit) !== -1 || pert === crit || pert.indexOf(crit) !== -1 || gr.indexOf(crit) !== -1) {
      res.push({
        cedula: cc,
        grado: datos[i][iGr],
        funcionario: datos[i][iNom],
        dependencia: datos[i][iDep],
        pert: datos[i][iPert],
        turno: datos[i][iTur]
      });
    }
  }
  return respuestaOK_('Resultados de búsqueda', res);
}

function obtenerFichaFuncionario(token, cedula) {
  validarSesion_(token);
  const hoja = obtenerHoja_(HOJA_BASE);
  const datos = hoja.getDataRange().getValues();
  const cc = normalizarCedula_(cedula);

  const enc = datos[0].map(v => normalizarTexto_(v));
  const iCc = enc.indexOf('CEDULA') >= 0 ? enc.indexOf('CEDULA') : 0;
  const iNiv = enc.indexOf('NIV') >= 0 ? enc.indexOf('NIV') : 1;
  const iGr = enc.indexOf('GR') >= 0 ? enc.indexOf('GR') : 2;
  const iNom = enc.indexOf('FUNCIONARIO') >= 0 ? enc.indexOf('FUNCIONARIO') : 3;
  const iDep = enc.indexOf('DEPENDENCIA') >= 0 ? enc.indexOf('DEPENDENCIA') : 4;
  const iPert = enc.indexOf('PERT') >= 0 ? enc.indexOf('PERT') : 6;
  const iTur = enc.indexOf('TURNO') >= 0 ? enc.indexOf('TURNO') : 7;

  for (let i = 1; i < datos.length; i++) {
    if (normalizarCedula_(datos[i][iCc]) === cc) {
      return {
        estado: true,
        funcionario: {
          cedula: cc,
          nivel: datos[i][iNiv],
          grado: datos[i][iGr],
          funcionario: datos[i][iNom],
          dependencia: datos[i][iDep],
          pert: datos[i][iPert],
          turno: datos[i][iTur]
        }
      };
    }
  }
  return respuestaError_('Funcionario no encontrado.');
}

function obtenerHistorialFuncionario(token, cedula) {
  validarSesion_(token);
  const hoja = obtenerHoja_(HOJA_NOVEDADES);
  const datos = hoja.getDataRange().getValues();
  const cc = normalizarCedula_(cedula);
  const hist = [];

  for (let i = 1; i < datos.length; i++) {
    if (normalizarCedula_(datos[i][2]) === cc) {
      hist.push({ GR: datos[i][0], NOVEDAD: datos[i][4], Dias: datos[i][6], 'Fecha INICIAL': datos[i][7], 'Fecha PRESENTACION': datos[i][8] });
    }
  }
  return { estado: true, historial: hist };
}

/* =====================================================
   CONFIGURACIÓN DE TIPOS DE NOVEDAD POR USUARIO
   ===================================================== */

const TIPOS_NOVEDAD_POR_USUARIO = {
  'SGSST_JESEP': ['EXCUSA MEDICA', 'RESTRICCIONES MEDICA', 'LICIENCIA DE MATERNIDAD'],
  'VAC_JESEP': ['PLAN VACACIONAL', 'VACACIONES EXTRAORDINARIAS', 'PLAN REDUCCION', 'VACACIONES DE RETIRO'],
  'PAS_JESEP': ['COMISION DE ESTUDIO', 'COMISION DE SERVICIO', 'LICENCIA DE PATERNIDAD', 'LICENCIA DE LUTO'],
  'CIT_JESEP': ['SUSPENCION', 'CITACION JUDICIAL (PERMISO)'],
  'HIS_JESEP': ['RETIROS 3 MESES DE ALTA', 'ELIMINAR USURIO POR RETIRO'],
  'PRO_JESEP': ['CURSO DE ASCENSO ESPOL','CURSO DE ASCENSO ESJIM'],
  'GH_JESEP': ['CAMBIO DE TURNO', 'HORARIO FLEXIBLE'],
  'CAP_JESEP': ['CURSO MADATORIO', 'CAPACITACION']
};

function obtenerTiposNovedadPermitidos(token) {
  const sesion = validarSesion_(token);
  const usr = String(sesion.usuario).toUpperCase();
  const tipos = TIPOS_NOVEDAD_POR_USUARIO[usr] || ['SERVICIO', 'PERMISO', 'OTRA NOVEDAD'];
  return respuestaOK_('Tipos asignados', { tipos: tipos });
}

function registrarNovedadPorUsuario(token, datos) {
  const sesion = validarSesion_(token);
  const tipo = datos.novedad;

  if (tipo === 'CAMBIO DE TURNO') {
    const hojaBase = obtenerHoja_(HOJA_BASE);
    const datosBase = hojaBase.getDataRange().getValues();
    const cc = normalizarCedula_(datos.cc || datos.cedula);
    const nuevoTurno = datos.nuevoTurno;

    const enc = datosBase[0].map(v => normalizarTexto_(v));
    const iCc = enc.indexOf('CEDULA') >= 0 ? enc.indexOf('CEDULA') : 0;
    const iTur = enc.indexOf('TURNO') >= 0 ? enc.indexOf('TURNO') : 7;

    for (let i = 1; i < datosBase.length; i++) {
      if (normalizarCedula_(datosBase[i][iCc]) === cc) {
        hojaBase.getRange(i + 1, iTur + 1).setValue(nuevoTurno);
        return respuestaOK_('Cambio de turno registrado en LISTADO_BASE correctamente.');
      }
    }
    return respuestaError_('Funcionario no encontrado.');
  }

  return registrarNovedad(token, datos);
}

/* =====================================================
   REGISTRAR NOVEDAD (COLUMNA K FORMATO Y FECHA DD/MM/YYYY)
   ===================================================== */

function registrarNovedad(token, datos) {
  validarSesion_(token);
  
  const cc = normalizarCedula_(datos.cc || datos.cedula || datos.CEDULA);
  if (!cc) return respuestaError_('No se recibió número de cédula válido.');

  const hojaBase = obtenerHoja_(HOJA_BASE);
  const datosBase = hojaBase.getDataRange().getValues();
  if (datosBase.length < 2) return respuestaError_('HOJA LISTADO_BASE vacía.');

  const encBase = datosBase[0].map(v => normalizarTexto_(v));
  const iCc = encBase.indexOf('CEDULA') >= 0 ? encBase.indexOf('CEDULA') : 0;
  const iGr = encBase.indexOf('GR') >= 0 ? encBase.indexOf('GR') : 2;
  const iNom = encBase.indexOf('FUNCIONARIO') >= 0 ? encBase.indexOf('FUNCIONARIO') : 3;
  const iDep = encBase.indexOf('DEPENDENCIA') >= 0 ? encBase.indexOf('DEPENDENCIA') : 4;
  const iNiv = encBase.indexOf('NIV') >= 0 ? encBase.indexOf('NIV') : 1;
  const iTur = encBase.indexOf('TURNO') >= 0 ? encBase.indexOf('TURNO') : 7;

  let gr = datos.grado || datos.GR || '';
  let funcionario = datos.funcionario || datos.FUNCIONARIO || '';
  let dependencia = datos.dependencia || '';
  let nivel = datos.nivel || '';
  let turno = datos.turno || '';

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

  if (!funcionario) {
    return respuestaError_('El funcionario con CC ' + cc + ' no existe en LISTADO_BASE.');
  }

  const tipoNovedadColE = datos.novedad || '';

  let descripcionColF = '';
  const reub = datos.descripcion || '';
  const detalleExtra = datos.nombreNovedad || '';

  if (reub && detalleExtra) {
    descripcionColF = reub + ' "' + detalleExtra + '"';
  } else if (reub) {
    descripcionColF = reub;
  } else {
    descripcionColF = detalleExtra;
  }

  // Formatear fechas a día/mes/año (DD/MM/YYYY)
  const fechaIniFormateada = formatearFechaLatina_(datos.fechaInicial);
  const fechaPresFormateada = formatearFechaLatina_(datos.fechaPresentacion);

  const hojaNov = obtenerHoja_(HOJA_NOVEDADES);
  
  const nuevaFila = [
    gr,                     // Columna A: GR
    funcionario,            // Columna B: APELLIDOS Y NOMBRES
    cc,                     // Columna C: CC
    dependencia,            // Columna D: DEPENDENCIA
    tipoNovedadColE,        // Columna E: NOVEDAD
    descripcionColF,        // Columna F: DESCRIPCION
    datos.dias || '',       // Columna G: Dias
    fechaIniFormateada,     // Columna H: Fecha INICIAL (DD/MM/YYYY)
    fechaPresFormateada,    // Columna I: Fecha PRESENTACION (DD/MM/YYYY)
    '',                     // Columna J: Observacion
    nivel,                  // Columna K: NIV (Ajustado aquí)
    '',                     // Columna L: Vacía
    turno,                  // Columna M: Turno
    '',                     // Columna N: Placa_Chip
    ''                      // Columna O: TEXTO
  ];

  hojaNov.appendRow(nuevaFila);
  return respuestaOK_('Novedad guardada exitosamente.');
}

function previsualizarSoloNovedades(token, filtro) {
  validarSesion_(token);
  const hojaNov = obtenerHoja_(HOJA_NOVEDADES);
  const datosNov = hojaNov.getDataRange().getValues();
  let filasHTML = '';

  for (let i = 1; i < datosNov.length; i++) {
    filasHTML += '<tr><td>' + datosNov[i][0] + '</td><td>' + datosNov[i][1] + '</td><td>' + datosNov[i][2] + '</td><td>' + datosNov[i][4] + '</td><td>' + datosNov[i][5] + '</td><td>' + datosNov[i][6] + '</td><td>' + datosNov[i][7] + '</td></tr>';
  }

  const html = '<html><head><style>table{width:100%;border-collapse:collapse;}th,td{border:1px solid #ddd;padding:8px;font-size:12px;}th{background:#01592F;color:white;}</style></head><body>' +
    '<h3>Reporte de Novedades Registradas - Turno ' + filtro + '</h3>' +
    '<table><thead><tr><th>GR</th><th>FUNCIONARIO</th><th>CÉDULA</th><th>NOVEDAD</th><th>DETALLE</th><th>DÍAS</th><th>INICIO</th></tr></thead><tbody>' +
    (filasHTML || '<tr><td colspan="7">No hay novedades registradas.</td></tr>') +
    '</tbody></table></body></html>';

  return { estado: true, html: html };
}

/* =====================================================
   MÓDULO UBL_JESEP — GESTIÓN LISTADO_BASE
   ===================================================== */

function agregarFuncionarioUBL(token, datos) {
  const sesion = validarSesion_(token);
  if (String(sesion.usuario).toUpperCase() !== 'UBL_JESEP' && sesion.rol !== 'ADMINISTRADOR') {
    throw new Error('No autorizado.');
  }

  const hoja = obtenerHoja_(HOJA_BASE);
  const fechaNacFormateada = formatearFechaLatina_(datos.fechaNacimiento);

  const fila = [
    datos.cedula,           // A: CEDULA
    datos.niv,              // B: NIV
    datos.gr,               // C: GR
    datos.funcionario,      // D: FUNCIONARIO
    datos.dependencia,      // E: DEPENDENCIA
    datos.dependencia,      // F: DEP
    datos.pert,             // G: PERT
    datos.turno,            // H: TURNO
    '',                     // I: NOVEDAD
    '',                     // J: OBSERVACIONES
    datos.mes,              // K: mes
    datos.dia,              // L: dia
    fechaNacFormateada,     // M: FECHA_NACIMIENTO (DD/MM/YYYY)
    datos.correo,           // N: CORREO_ELECTRONICO
    datos.sexo,             // O: SEXO
    datos.estadoCivil,      // P: ESTADO_CIVIL
    datos.situacionLaboral, // Q: SITUACION_LABORAL
    '', '', '', '', '', '', '', '', '', '', '', '', '', '', '',
    datos.comunicadoOficial // AG: COMUNICADO_OFICIAL
  ];

  hoja.appendRow(fila);
  return respuestaOK_('Funcionario agregado exitosamente a LISTADO_BASE.');
}

function eliminarFuncionarioUBL(token, cedula) {
  const sesion = validarSesion_(token);
  if (String(sesion.usuario).toUpperCase() !== 'UBL_JESEP' && sesion.rol !== 'ADMINISTRADOR') {
    throw new Error('No autorizado.');
  }

  const hoja = obtenerHoja_(HOJA_BASE);
  const datos = hoja.getDataRange().getValues();
  const cc = normalizarCedula_(cedula);

  for (let i = 1; i < datos.length; i++) {
    if (normalizarCedula_(datos[i][0]) === cc) {
      hoja.deleteRow(i + 1);
      return respuestaOK_('Funcionario eliminado correctamente de LISTADO_BASE.');
    }
  }
  return respuestaError_('Funcionario no encontrado.');
}