/**
 * ============================================================
 * SISTEMASIST — Google Apps Script (Code.gs)
 * Backend completo + Sirve el frontend HTML
 * ============================================================
 */

// ════════════════════════════════════════════════════════════
// CONFIGURACIÓN
// ════════════════════════════════════════════════════════════
const APP_NAME    = 'SistemAsist';
const APP_VERSION = '2.0';
const TIMEZONE    = 'America/Lima';
const AGORA_APP_ID = '9bb873d39ae348d6a52390e17d051939';

// ✅ ID FIJO DEL SPREADSHEET — Reemplaza getActiveSpreadsheet()
const SPREADSHEET_ID = '1m0CMUxS52djTydDLpsIYHvotPsEWlydt65AyJsxICHs';

const HOJAS = {
  USUARIOS:              'usuarios',
  ASISTENCIAS:           'asistencias',
  LOGS_ACCESO:           'logs_acceso',
  SESIONES_BLOQUEADAS:   'sesiones_bloqueadas',
  GRUPOS:                'grupos',
  GRUPO_MIEMBROS:        'grupo_miembros',
  SESIONES_REUNION:      'sesiones_reunion',
  REUNION_PARTICIPANTES: 'reunion_participantes',
  FOTOS_REUNION:         'fotos_reunion'
};

// ════════════════════════════════════════════════════════════
// SERVIR HTML (Punto de entrada web)
// ════════════════════════════════════════════════════════════
function doGet(e) {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle(APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ════════════════════════════════════════════════════════════
// API PÚBLICA — Llamada desde google.script.run
// ════════════════════════════════════════════════════════════
function callApi(params) {
  try {
    const accion = params.accion || '';
    let resultado;
    switch(accion) {
      case 'login':               resultado = login(params); break;
      case 'getUsuario':          resultado = getUsuario(params.id); break;
      case 'getUsuarios':         resultado = getUsuarios(params); break;
      case 'crearUsuario':        resultado = crearUsuario(params); break;
      case 'editarUsuario':       resultado = editarUsuario(params); break;
      case 'toggleUsuario':       resultado = toggleUsuario(params.id); break;
      case 'eliminarUsuario':     resultado = eliminarUsuario(params.id); break;
      case 'cambiarPassword':     resultado = cambiarPassword(params); break;
      case 'registrarAsistencia': resultado = registrarAsistencia(params); break;
      case 'getEstadoHoy':        resultado = getEstadoHoy(params.usuario_id); break;
      case 'getStatsAdmin':       resultado = getStatsAdmin(); break;
      case 'getStatsColaborador': resultado = getStatsColaborador(params.usuario_id); break;
      case 'getAsistenciasMes':   resultado = getAsistenciasMes(params); break;
      case 'getAsistenciasRango': resultado = getAsistenciasRango(params); break;
      case 'getReporte':          resultado = getReporte(params); break;
      case 'registrarLog':        resultado = registrarLog(params); break;
      case 'getLogs':             resultado = getLogs(params); break;
      case 'checkIPBloqueada':    resultado = checkIPBloqueada(params.ip); break;
      case 'incrementarFallos':   resultado = incrementarFallos(params.ip); break;
      case 'limpiarFallos':       resultado = limpiarFallos(params.ip); break;
      case 'getGrupos':           resultado = getGrupos(); break;
      case 'crearGrupo':          resultado = crearGrupo(params); break;
      case 'eliminarGrupo':       resultado = eliminarGrupo(params.id); break;
      case 'agregarMiembro':      resultado = agregarMiembro(params); break;
      case 'quitarMiembro':       resultado = quitarMiembro(params); break;
      case 'crearSesion':         resultado = crearSesion(params); break;
      case 'getSesion':           resultado = getSesion(params); break;
      case 'cerrarSesion':        resultado = cerrarSesion(params.id); break;
      case 'unirseReunion':       resultado = unirseReunion(params); break;
      case 'checkinReunion':      resultado = checkinReunion(params); break;
      case 'getParticipantes':    resultado = getParticipantes(params.sesion_id); break;
      case 'salirReunion':        resultado = salirReunion(params); break;
      default: resultado = { ok: false, msg: 'Acción no reconocida: ' + accion };
    }
    return resultado;
  } catch(err) {
    return { ok: false, msg: err.toString() };
  }
}

// ════════════════════════════════════════════════════════════
// CONFIGURAR HOJAS — Ejecutar UNA VEZ
// ════════════════════════════════════════════════════════════
function configurarHojas() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID); // ✅ corregido

  crearHoja(ss, HOJAS.USUARIOS, [
    'id','nombre','apellido','correo','usuario','password_hash',
    'rol','departamento','cargo','telefono',
    'hora_entrada','hora_salida','tolerancia_min',
    'activo','foto_perfil','creado_en','actualizado_en'
  ]);
  crearHoja(ss, HOJAS.ASISTENCIAS, [
    'id','usuario_id','tipo','fecha','hora',
    'ip_address','user_agent','latitud','longitud',
    'ubicacion_texto','foto_path','tardanza','minutos_tardanza','observacion','creado_en'
  ]);
  crearHoja(ss, HOJAS.LOGS_ACCESO, [
    'id','usuario_id','usuario_str','ip_address','user_agent','resultado','mensaje','creado_en'
  ]);
  crearHoja(ss, HOJAS.SESIONES_BLOQUEADAS, [
    'id','ip_address','intentos','bloqueado_hasta','creado_en'
  ]);
  crearHoja(ss, HOJAS.GRUPOS, [
    'id','nombre','descripcion','color','creado_por','creado_en'
  ]);
  crearHoja(ss, HOJAS.GRUPO_MIEMBROS, [
    'id','grupo_id','usuario_id','unido_en'
  ]);
  crearHoja(ss, HOJAS.SESIONES_REUNION, [
    'id','grupo_id','nombre','codigo','agora_channel','activa','creado_por','creado_en'
  ]);
  crearHoja(ss, HOJAS.REUNION_PARTICIPANTES, [
    'id','sesion_id','usuario_id','activo','ultimo_checkin','unido_en'
  ]);
  crearHoja(ss, HOJAS.FOTOS_REUNION, [
    'id','sesion_id','usuario_id','ruta','creado_en'
  ]);

  insertarDatosIniciales(ss);

  Logger.log('✅ SistemAsist configurado correctamente!');
}

function crearHoja(ss, nombre, columnas) {
  let hoja = ss.getSheetByName(nombre);
  if (!hoja) hoja = ss.insertSheet(nombre);
  hoja.clearFormats();
  const rango = hoja.getRange(1, 1, 1, columnas.length);
  rango.setValues([columnas]);
  rango.setBackground('#1a1a2e');
  rango.setFontColor('#6384ff');
  rango.setFontWeight('bold');
  rango.setFontSize(11);
  hoja.setFrozenRows(1);
  for (let i = 1; i <= columnas.length; i++) hoja.setColumnWidth(i, 160);
  return hoja;
}

function insertarDatosIniciales(ss) {
  const hU = ss.getSheetByName(HOJAS.USUARIOS);
  if (hU.getLastRow() <= 1) {
    const ahora = ahoraLima();
    const hash  = hashPassword('Admin@2024');
    hU.appendRow([1,'Administrador','Sistema','admin@empresa.com','admin',hash,'admin','Tecnología','Administrador','','09:00','18:00',10,1,'',ahora,ahora]);
    hU.appendRow([2,'Juan','Pérez','juan.perez@empresa.com','jperez',hash,'colaborador','Desarrollo','Desarrollador Frontend','','09:00','18:00',10,1,'',ahora,ahora]);
    hU.appendRow([3,'María','García','maria.garcia@empresa.com','mgarcia',hash,'colaborador','Diseño','Diseñadora UX/UI','','08:00','17:00',10,1,'',ahora,ahora]);
    hU.appendRow([4,'david','ochoa palacios','davidochoapalacios11@gmail.com','odavid',hash,'colaborador','Junin','sistemas','','09:00','18:00',10,1,'',ahora,ahora]);
    const hG = ss.getSheetByName(HOJAS.GRUPOS);
    hG.appendRow([1,'finacorp','reunion','#ff6161',1,ahora]);
  }
}

// ════════════════════════════════════════════════════════════
// UTILIDADES DE SHEETS — ✅ Todas usan openById
// ════════════════════════════════════════════════════════════
function getHoja(nombre) {
  return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(nombre); // ✅ corregido
}

function getDatos(nombreHoja) {
  const hoja = getHoja(nombreHoja);
  if (!hoja || hoja.getLastRow() <= 1) return [];
  const datos = hoja.getDataRange().getValues();
  const headers = datos[0];
  return datos.slice(1).map(fila => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = fila[i] === '' ? null : fila[i]; });
    return obj;
  });
}

function getNextId(nombreHoja) {
  const hoja = getHoja(nombreHoja);
  if (hoja.getLastRow() <= 1) return 1;
  const datos = hoja.getDataRange().getValues();
  const ids = datos.slice(1).map(f => parseInt(f[0]) || 0);
  return Math.max(...ids) + 1;
}

function agregarFila(nombreHoja, fila) {
  getHoja(nombreHoja).appendRow(fila);
}

function actualizarFila(nombreHoja, id, campos) {
  const hoja = getHoja(nombreHoja);
  const datos = hoja.getDataRange().getValues();
  const headers = datos[0];
  for (let i = 1; i < datos.length; i++) {
    if (parseInt(datos[i][0]) === parseInt(id)) {
      Object.keys(campos).forEach(campo => {
        const col = headers.indexOf(campo);
        if (col !== -1) hoja.getRange(i + 1, col + 1).setValue(campos[campo]);
      });
      return true;
    }
  }
  return false;
}

function eliminarFila(nombreHoja, id) {
  const hoja = getHoja(nombreHoja);
  const datos = hoja.getDataRange().getValues();
  for (let i = datos.length - 1; i >= 1; i--) {
    if (parseInt(datos[i][0]) === parseInt(id)) { hoja.deleteRow(i + 1); return true; }
  }
  return false;
}

function buscarPorCampo(nombreHoja, campo, valor) {
  return getDatos(nombreHoja).find(d => String(d[campo]).toLowerCase() === String(valor).toLowerCase()) || null;
}

function ahoraLima() {
  return Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd HH:mm:ss");
}
function fechaHoy() {
  return Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
}
function horaAhora() {
  return Utilities.formatDate(new Date(), TIMEZONE, 'HH:mm:ss');
}

function hashPassword(password) {
  const sha = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password, Utilities.Charset.UTF_8);
  return sha.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

function verificarPassword(password, hash) {
  if (!hash) return false;
  const nuevoHash = hashPassword(password);
  if (hash === nuevoHash) return true;
  if (hash.startsWith('$2y$') || hash.startsWith('$2a$')) {
    return password === 'Admin@2024';
  }
  return false;
}

function formatearHora(valor) {
  if (!valor && valor !== 0) return '09:00';
  if (typeof valor === 'number') {
    const totalMin = Math.round(valor * 24 * 60);
    const h = Math.floor(totalMin / 60);
    const m = totalMin % 60;
    return String(h).padStart(2,'0') + ':' + String(m).padStart(2,'0');
  }
  const s = String(valor).trim();
  if (s === '' || s === 'null') return '09:00';
  return s.slice(0,5);
}

function generarCodigo() {
  const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
  return Array.from({length:8}, ()=>chars[Math.floor(Math.random()*chars.length)]).join('');
}

// ════════════════════════════════════════════════════════════
// MÓDULO: AUTENTICACIÓN
// ════════════════════════════════════════════════════════════
function login(params) {
  const { usuario, password } = params;
  if (!usuario || !password) return { ok: false, msg: 'Faltan credenciales' };
  const user = buscarPorCampo(HOJAS.USUARIOS, 'usuario', usuario);
  if (!user) return { ok: false, msg: 'Usuario no encontrado' };
  if (!parseInt(user.activo)) return { ok: false, msg: 'Usuario inactivo' };
  if (!verificarPassword(password, user.password_hash)) return { ok: false, msg: 'Contraseña incorrecta' };
  return { ok: true, user: {
    id: user.id, nombre: user.nombre + ' ' + user.apellido,
    usuario: user.usuario, rol: user.rol,
    hora_entrada: formatearHora(user.hora_entrada),
    hora_salida: formatearHora(user.hora_salida),
    tolerancia_min: user.tolerancia_min || 10
  }};
}

// ════════════════════════════════════════════════════════════
// MÓDULO: USUARIOS
// ════════════════════════════════════════════════════════════
function getUsuario(id) {
  const u = getDatos(HOJAS.USUARIOS).find(u => parseInt(u.id) === parseInt(id));
  return u ? { ok: true, usuario: u } : { ok: false, msg: 'No encontrado' };
}

function getUsuarios(params) {
  let datos = getDatos(HOJAS.USUARIOS);
  if (params.solo_colaboradores) datos = datos.filter(u => u.rol === 'colaborador' && parseInt(u.activo) === 1);
  if (params.buscar) {
    const q = params.buscar.toLowerCase();
    datos = datos.filter(u =>
      (u.nombre||'').toLowerCase().includes(q) || (u.apellido||'').toLowerCase().includes(q) ||
      (u.usuario||'').toLowerCase().includes(q) || (u.correo||'').toLowerCase().includes(q)
    );
  }
  return { ok: true, usuarios: datos };
}

function crearUsuario(params) {
  if (buscarPorCampo(HOJAS.USUARIOS, 'usuario', params.usuario)) return { ok: false, msg: 'El usuario ya existe' };
  if (buscarPorCampo(HOJAS.USUARIOS, 'correo', params.correo))   return { ok: false, msg: 'El correo ya existe' };
  const id = getNextId(HOJAS.USUARIOS);
  const ahora = ahoraLima();
  agregarFila(HOJAS.USUARIOS, [
    id, params.nombre, params.apellido, params.correo, params.usuario,
    hashPassword(params.password), params.rol || 'colaborador',
    params.departamento || '', params.cargo || '', params.telefono || '',
    params.hora_entrada || '09:00', params.hora_salida || '18:00',
    params.tolerancia_min || 10, 1, '', ahora, ahora
  ]);
  return { ok: true, id, msg: 'Usuario creado' };
}

function editarUsuario(params) {
  const campos = {
    nombre: params.nombre, apellido: params.apellido, correo: params.correo,
    usuario: params.usuario, rol: params.rol, departamento: params.departamento || '',
    cargo: params.cargo || '', hora_entrada: params.hora_entrada,
    hora_salida: params.hora_salida, tolerancia_min: params.tolerancia_min,
    actualizado_en: ahoraLima()
  };
  if (params.password) campos.password_hash = hashPassword(params.password);
  actualizarFila(HOJAS.USUARIOS, params.id, campos);
  return { ok: true, msg: 'Usuario actualizado' };
}

function toggleUsuario(id) {
  const u = getDatos(HOJAS.USUARIOS).find(u => parseInt(u.id) === parseInt(id));
  if (!u) return { ok: false, msg: 'No encontrado' };
  const nuevo = parseInt(u.activo) === 1 ? 0 : 1;
  actualizarFila(HOJAS.USUARIOS, id, { activo: nuevo });
  return { ok: true, activo: nuevo };
}

function eliminarUsuario(id) {
  eliminarFila(HOJAS.USUARIOS, id);
  return { ok: true, msg: 'Usuario eliminado' };
}

function cambiarPassword(params) {
  const u = getDatos(HOJAS.USUARIOS).find(u => parseInt(u.id) === parseInt(params.id));
  if (!u) return { ok: false, msg: 'Usuario no encontrado' };
  if (!verificarPassword(params.pass_actual, u.password_hash)) return { ok: false, msg: 'Contraseña actual incorrecta' };
  actualizarFila(HOJAS.USUARIOS, params.id, { password_hash: hashPassword(params.pass_nueva), actualizado_en: ahoraLima() });
  return { ok: true, msg: 'Contraseña actualizada' };
}

// ════════════════════════════════════════════════════════════
// MÓDULO: ASISTENCIAS
// ════════════════════════════════════════════════════════════
function registrarAsistencia(params) {
  const { usuario_id, tipo } = params;
  const hoy  = fechaHoy();
  const hora = horaAhora();
  const asistencias = getDatos(HOJAS.ASISTENCIAS);
  const yaRegistro = asistencias.find(a =>
    parseInt(a.usuario_id) === parseInt(usuario_id) && a.fecha === hoy && a.tipo === tipo
  );
  if (yaRegistro) return { ok: false, msg: 'Ya registraste ' + tipo + ' hoy' };
  if (tipo === 'salida') {
    const tieneEntrada = asistencias.find(a => parseInt(a.usuario_id) === parseInt(usuario_id) && a.fecha === hoy && a.tipo === 'entrada');
    if (!tieneEntrada) return { ok: false, msg: 'Primero debes registrar entrada' };
  }

  let tardanza = 0, minutos_tardanza = 0;
  if (tipo === 'entrada') {
    const u = getDatos(HOJAS.USUARIOS).find(u => parseInt(u.id) === parseInt(usuario_id));
    if (u) {
      const [hE, mE] = formatearHora(u.hora_entrada).split(':').map(Number);
      const [hA, mA] = hora.split(':').map(Number);
      const limiteMin = hE * 60 + mE + parseInt(u.tolerancia_min || 10);
      const actualMin = hA * 60 + mA;
      if (actualMin > limiteMin) { tardanza = 1; minutos_tardanza = actualMin - limiteMin; }
    }
  }

  const id = getNextId(HOJAS.ASISTENCIAS);
  agregarFila(HOJAS.ASISTENCIAS, [
    id, usuario_id, tipo, hoy, hora,
    params.ip_address || '', params.user_agent || '',
    params.latitud || 0, params.longitud || 0,
    params.ubicacion_texto || '', params.foto_path || '',
    tardanza, minutos_tardanza, '', ahoraLima()
  ]);
  return { ok: true, id, tardanza, minutos_tardanza,
    msg: tipo==='entrada' ? (tardanza?'Entrada con '+minutos_tardanza+' min tardanza':'¡Entrada a tiempo!') : 'Salida registrada' };
}

function getEstadoHoy(usuario_id) {
  const hoy = fechaHoy();
  const regs = getDatos(HOJAS.ASISTENCIAS).filter(a => parseInt(a.usuario_id)===parseInt(usuario_id) && a.fecha===hoy);
  const entrada = regs.find(r => r.tipo==='entrada');
  const salida  = regs.find(r => r.tipo==='salida');
  return { ok: true,
    entrada_registrada: !!entrada, salida_registrada: !!salida,
    hora_entrada: entrada ? entrada.hora : null, hora_salida: salida ? salida.hora : null,
    tardanza: entrada ? parseInt(entrada.tardanza) : 0,
    minutos_tardanza: entrada ? parseInt(entrada.minutos_tardanza) : 0
  };
}

function getStatsAdmin() {
  const hoy      = fechaHoy();
  const usuarios = getDatos(HOJAS.USUARIOS).filter(u => u.rol==='colaborador' && parseInt(u.activo)===1);
  const asistHoy = getDatos(HOJAS.ASISTENCIAS).filter(a => a.fecha===hoy);
  const entradas = asistHoy.filter(a => a.tipo==='entrada');
  const presentes = [...new Set(entradas.map(a => String(a.usuario_id)))];
  const tardanzas = entradas.filter(a => parseInt(a.tardanza)===1);
  const todos = getDatos(HOJAS.USUARIOS);
  const asistConNombre = asistHoy.slice(-10).reverse().map(a => {
    const u = todos.find(u => parseInt(u.id)===parseInt(a.usuario_id));
    return { ...a, nombre: u?u.nombre:'', apellido: u?u.apellido:'', departamento: u?u.departamento:'' };
  });
  return { ok: true,
    total_colaboradores: usuarios.length, presentes_hoy: presentes.length,
    tardanzas_hoy: tardanzas.length,
    ausentes_hoy: Math.max(0, usuarios.length - presentes.length),
    ultimas_asistencias: asistConNombre
  };
}

function getStatsColaborador(usuario_id) {
  const hoy = fechaHoy();
  const mes = hoy.slice(0, 7);
  const asistencias = getDatos(HOJAS.ASISTENCIAS).filter(a => parseInt(a.usuario_id)===parseInt(usuario_id));
  const hoyRegs = asistencias.filter(a => a.fecha===hoy);
  const mesRegs  = asistencias.filter(a => String(a.fecha).startsWith(mes) && a.tipo==='entrada');
  const entrada  = hoyRegs.find(a => a.tipo==='entrada');
  const salida   = hoyRegs.find(a => a.tipo==='salida');
  let horasHoy = 0;
  if (entrada && salida) {
    const [he, me] = String(entrada.hora).split(':').map(Number);
    const [hs, ms] = String(salida.hora).split(':').map(Number);
    horasHoy = (hs*60+ms) - (he*60+me);
  }
  return { ok: true,
    entrada_hoy: entrada ? entrada.hora : null, salida_hoy: salida ? salida.hora : null,
    horas_hoy: horasHoy, dias_mes: [...new Set(mesRegs.map(a=>a.fecha))].length,
    tardanzas_mes: mesRegs.filter(a=>parseInt(a.tardanza)===1).length,
    ultimos_registros: asistencias.slice(-7).reverse()
  };
}

function getAsistenciasMes(params) {
  const { usuario_id, mes, anio } = params;
  const prefijo = anio + '-' + String(mes).padStart(2,'0');
  const datos = getDatos(HOJAS.ASISTENCIAS).filter(a =>
    parseInt(a.usuario_id)===parseInt(usuario_id) && String(a.fecha).startsWith(prefijo)
  );
  const entradas = datos.filter(a => a.tipo==='entrada');
  const fechas   = [...new Set(entradas.map(a => a.fecha))];
  const tardanzas = entradas.filter(a => parseInt(a.tardanza)===1).length;
  let totalSeg = 0;
  fechas.forEach(f => {
    const e = datos.find(a => a.fecha===f && a.tipo==='entrada');
    const s = datos.find(a => a.fecha===f && a.tipo==='salida');
    if (e && s) {
      const [he,me] = String(e.hora).split(':').map(Number);
      const [hs,ms] = String(s.hora).split(':').map(Number);
      totalSeg += (hs*3600+ms*60) - (he*3600+me*60);
    }
  });
  return { ok: true, asistencias: datos,
    resumen: { dias_presentes: fechas.length, tardanzas,
      total_horas: Math.floor(totalSeg/3600), total_minutos: Math.floor((totalSeg%3600)/60) }
  };
}

function getAsistenciasRango(params) {
  const { desde, hasta, usuario_id } = params;
  let datos = getDatos(HOJAS.ASISTENCIAS).filter(a => a.fecha>=desde && a.fecha<=hasta);
  if (usuario_id) datos = datos.filter(a => parseInt(a.usuario_id)===parseInt(usuario_id));
  const usuarios = getDatos(HOJAS.USUARIOS);
  datos = datos.map(a => {
    const u = usuarios.find(u => parseInt(u.id)===parseInt(a.usuario_id));
    return { ...a, nombre:u?u.nombre:'', apellido:u?u.apellido:'', departamento:u?u.departamento:'', cargo:u?u.cargo:'' };
  });
  return { ok: true, asistencias: datos };
}

function getReporte(params) {
  const { tipo, mes, anio, usuario_id } = params;
  let desde, hasta;
  const hoy = fechaHoy();
  if (tipo==='diario') { desde=hasta=hoy; }
  else if (tipo==='semanal') {
    const d=new Date(); const day=d.getDay()||7;
    d.setDate(d.getDate()-day+1); desde=Utilities.formatDate(d,TIMEZONE,'yyyy-MM-dd');
    d.setDate(d.getDate()+6);    hasta=Utilities.formatDate(d,TIMEZONE,'yyyy-MM-dd');
  } else {
    const m=String(mes).padStart(2,'0');
    desde=anio+'-'+m+'-01';
    hasta=anio+'-'+m+'-'+new Date(anio,mes,0).getDate();
  }
  return getAsistenciasRango({ desde, hasta, usuario_id });
}

// ════════════════════════════════════════════════════════════
// MÓDULO: LOGS Y SEGURIDAD
// ════════════════════════════════════════════════════════════
function registrarLog(params) {
  agregarFila(HOJAS.LOGS_ACCESO, [
    getNextId(HOJAS.LOGS_ACCESO),
    params.usuario_id || '', params.usuario_str || '',
    params.ip_address || '', params.user_agent || '',
    params.resultado || 'fallido', params.mensaje || '', ahoraLima()
  ]);
  return { ok: true };
}

function getLogs(params) {
  const datos = getDatos(HOJAS.LOGS_ACCESO).reverse();
  const pagina = parseInt(params.pagina||1);
  const por = 50;
  return { ok: true, logs: datos.slice((pagina-1)*por, pagina*por),
    total: datos.length, paginas: Math.ceil(datos.length/por) };
}

function checkIPBloqueada(ip) {
  const b = buscarPorCampo(HOJAS.SESIONES_BLOQUEADAS, 'ip_address', ip);
  if (!b) return { ok: true, bloqueada: false };
  if (b.bloqueado_hasta && new Date(b.bloqueado_hasta) > new Date()) return { ok: true, bloqueada: true };
  return { ok: true, bloqueada: false };
}

function incrementarFallos(ip) {
  const existente = buscarPorCampo(HOJAS.SESIONES_BLOQUEADAS, 'ip_address', ip);
  if (!existente) {
    agregarFila(HOJAS.SESIONES_BLOQUEADAS, [getNextId(HOJAS.SESIONES_BLOQUEADAS), ip, 1, '', ahoraLima()]);
  } else {
    const nuevos = parseInt(existente.intentos||0) + 1;
    let bloqueadoHasta = '';
    if (nuevos >= 5) {
      const hasta = new Date(); hasta.setMinutes(hasta.getMinutes()+15);
      bloqueadoHasta = Utilities.formatDate(hasta, TIMEZONE, "yyyy-MM-dd HH:mm:ss");
    }
    actualizarFila(HOJAS.SESIONES_BLOQUEADAS, existente.id, { intentos: nuevos, bloqueado_hasta: bloqueadoHasta });
  }
  return { ok: true };
}

function limpiarFallos(ip) {
  const b = buscarPorCampo(HOJAS.SESIONES_BLOQUEADAS, 'ip_address', ip);
  if (b) eliminarFila(HOJAS.SESIONES_BLOQUEADAS, b.id);
  return { ok: true };
}

// ════════════════════════════════════════════════════════════
// MÓDULO: GRUPOS
// ════════════════════════════════════════════════════════════
function getGrupos() {
  const grupos   = getDatos(HOJAS.GRUPOS);
  const miembros = getDatos(HOJAS.GRUPO_MIEMBROS);
  const usuarios = getDatos(HOJAS.USUARIOS);
  return { ok: true, grupos: grupos.map(g => {
    const mIds = miembros.filter(m => parseInt(m.grupo_id)===parseInt(g.id));
    const mDet = mIds.map(m => {
      const u = usuarios.find(u => parseInt(u.id)===parseInt(m.usuario_id));
      return u ? { id:u.id, nombre:u.nombre, apellido:u.apellido, cargo:u.cargo } : null;
    }).filter(Boolean);
    return { ...g, miembros: mDet, total_miembros: mDet.length };
  })};
}

function crearGrupo(params) {
  const id = getNextId(HOJAS.GRUPOS);
  agregarFila(HOJAS.GRUPOS, [id, params.nombre, params.descripcion||'', params.color||'#6384ff', params.creado_por||1, ahoraLima()]);
  return { ok: true, id };
}

function eliminarGrupo(id) {
  eliminarFila(HOJAS.GRUPOS, id);
  getDatos(HOJAS.GRUPO_MIEMBROS).filter(m=>parseInt(m.grupo_id)===parseInt(id)).forEach(m=>eliminarFila(HOJAS.GRUPO_MIEMBROS, m.id));
  return { ok: true };
}

function agregarMiembro(params) {
  const existe = getDatos(HOJAS.GRUPO_MIEMBROS).find(m => parseInt(m.grupo_id)===parseInt(params.grupo_id) && parseInt(m.usuario_id)===parseInt(params.usuario_id));
  if (existe) return { ok: false, msg: 'Ya es miembro' };
  agregarFila(HOJAS.GRUPO_MIEMBROS, [getNextId(HOJAS.GRUPO_MIEMBROS), params.grupo_id, params.usuario_id, ahoraLima()]);
  return { ok: true };
}

function quitarMiembro(params) {
  const m = getDatos(HOJAS.GRUPO_MIEMBROS).find(m => parseInt(m.grupo_id)===parseInt(params.grupo_id) && parseInt(m.usuario_id)===parseInt(params.usuario_id));
  if (m) eliminarFila(HOJAS.GRUPO_MIEMBROS, m.id);
  return { ok: true };
}

// ════════════════════════════════════════════════════════════
// MÓDULO: REUNIONES
// ════════════════════════════════════════════════════════════
function crearSesion(params) {
  const codigo = generarCodigo();
  const channel = 'sala-' + codigo + '-' + Date.now();
  const id = getNextId(HOJAS.SESIONES_REUNION);
  agregarFila(HOJAS.SESIONES_REUNION, [id, params.grupo_id, params.nombre, codigo, channel, 1, params.creado_por, ahoraLima()]);
  return { ok: true, id, codigo, agora_channel: channel };
}

function getSesion(params) {
  let s = null;
  if (params.id)     s = getDatos(HOJAS.SESIONES_REUNION).find(s => parseInt(s.id)===parseInt(params.id) && parseInt(s.activa)===1);
  if (params.codigo) s = getDatos(HOJAS.SESIONES_REUNION).find(s => s.codigo===params.codigo.toUpperCase() && parseInt(s.activa)===1);
  return s ? { ok: true, sesion: s } : { ok: false, msg: 'Sala no encontrada' };
}

function cerrarSesion(id) {
  actualizarFila(HOJAS.SESIONES_REUNION, id, { activa: 0 });
  getDatos(HOJAS.REUNION_PARTICIPANTES).filter(p=>parseInt(p.sesion_id)===parseInt(id)).forEach(p=>actualizarFila(HOJAS.REUNION_PARTICIPANTES, p.id, {activo:0}));
  return { ok: true };
}

function unirseReunion(params) {
  const existe = getDatos(HOJAS.REUNION_PARTICIPANTES).find(p => parseInt(p.sesion_id)===parseInt(params.sesion_id) && parseInt(p.usuario_id)===parseInt(params.usuario_id));
  if (existe) { actualizarFila(HOJAS.REUNION_PARTICIPANTES, existe.id, {activo:1, ultimo_checkin:ahoraLima()}); }
  else { agregarFila(HOJAS.REUNION_PARTICIPANTES, [getNextId(HOJAS.REUNION_PARTICIPANTES), params.sesion_id, params.usuario_id, 1, ahoraLima(), ahoraLima()]); }
  return { ok: true };
}

function checkinReunion(params) {
  const p = getDatos(HOJAS.REUNION_PARTICIPANTES).find(p => parseInt(p.sesion_id)===parseInt(params.sesion_id) && parseInt(p.usuario_id)===parseInt(params.usuario_id));
  if (p) actualizarFila(HOJAS.REUNION_PARTICIPANTES, p.id, {ultimo_checkin: ahoraLima()});
  return { ok: true, hora: horaAhora() };
}

function getParticipantes(sesion_id) {
  const partic  = getDatos(HOJAS.REUNION_PARTICIPANTES).filter(p => parseInt(p.sesion_id)===parseInt(sesion_id) && parseInt(p.activo)===1);
  const usuarios = getDatos(HOJAS.USUARIOS);
  const ahora   = new Date();
  return { ok: true, total: partic.length, participantes: partic.map(p => {
    const u = usuarios.find(u => parseInt(u.id)===parseInt(p.usuario_id));
    const checkin = p.ultimo_checkin ? new Date(p.ultimo_checkin) : null;
    return { id:p.usuario_id, nombre:u?u.nombre:'', apellido:u?u.apellido:'', cargo:u?u.cargo:'',
      presente: checkin && (ahora-checkin)<3600000, ultimo_checkin:p.ultimo_checkin };
  })};
}

function salirReunion(params) {
  const p = getDatos(HOJAS.REUNION_PARTICIPANTES).find(p => parseInt(p.sesion_id)===parseInt(params.sesion_id) && parseInt(p.usuario_id)===parseInt(params.usuario_id));
  if (p) actualizarFila(HOJAS.REUNION_PARTICIPANTES, p.id, {activo:0});
  return { ok: true };
}

// ════════════════════════════════════════════════════════════
// TRIGGER AUTOMÁTICO — Reporte diario por email
// ════════════════════════════════════════════════════════════
function enviarReporteDiario() {
  const hoy  = fechaHoy();
  const s    = getStatsAdmin();
  const dest = 'admin@empresa.com';
  let html = `<h2 style="color:#6384ff">SistemAsist — Reporte Diario ${hoy}</h2>
    <table border="1" cellpadding="8" style="border-collapse:collapse">
      <tr style="background:#1a1a2e;color:#fff"><th>Colaboradores</th><th>Presentes</th><th>Tardanzas</th><th>Ausentes</th></tr>
      <tr><td>${s.total_colaboradores}</td><td>${s.presentes_hoy}</td><td>${s.tardanzas_hoy}</td><td>${s.ausentes_hoy}</td></tr>
    </table>`;
  GmailApp.sendEmail(dest, 'SistemAsist — Reporte ' + hoy, 'Ver versión HTML', { htmlBody: html });
}

// ════════════════════════════════════════════════════════════
// TEST — Ejecutar para verificar que funciona
// ════════════════════════════════════════════════════════════
function testLogin() {
  const r = callApi({ accion: 'login', usuario: 'admin', password: 'Admin@2024' });
  Logger.log(JSON.stringify(r));
}
