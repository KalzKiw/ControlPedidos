const SHEET_NAME = 'Hoja1'; // El nombre de la hoja dentro de "Pedidos"
const TABLE_NAME = 'Tabla1'; // El nombre de la tabla

// Función que maneja la solicitud GET para mostrar el formulario
function doGet() {
  return HtmlService.createHtmlOutputFromFile('formulario');
}

// Función para formatear fecha como texto DD/MM/YYYY
function formatearFecha(fecha) {
  const dia = String(fecha.getDate()).padStart(2, '0');
  const mes = String(fecha.getMonth() + 1).padStart(2, '0');
  const anio = fecha.getFullYear();
  return `${dia}/${mes}/${anio}`;
}

// Función para guardar el pedido en la hoja de cálculo
function guardarPedido(data) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);

  const fechaPedido = new Date(data.fechaPedido);
  const fechaEntrega = data.fechaEntrega ? new Date(data.fechaEntrega) : null;

  // Calcular la diferencia de días si "Fecha Entrega" está presente
  let diasDiferencia = "";
  if (fechaEntrega) {
    diasDiferencia = Math.ceil((fechaEntrega - fechaPedido) / (1000 * 60 * 60 * 24));
  }

  const incidenciaTexto = data.incidencia ? "Sí" : "No";
  const tipoIncidenciaTexto = data.incidencia && data.tipoIncidencia ? data.tipoIncidencia : "NO";

  // Determinar el estado
  const estado = fechaEntrega ? "Entregado" : "Pendiente";

  // Agregar los datos a la hoja
  hoja.appendRow([
    data.cliente || "", // Nueva columna para el cliente
    data.pedido || "",
    Utilities.formatDate(fechaPedido, Session.getScriptTimeZone(), "dd/MM/yyyy"),
    fechaEntrega ? Utilities.formatDate(fechaEntrega, Session.getScriptTimeZone(), "dd/MM/yyyy") : "",
    diasDiferencia,
    estado,
    incidenciaTexto,
    tipoIncidenciaTexto
  ]);
}

// Función para manejar el POST que recibe los datos y los guarda en la hoja de Google Sheets
function doPost(e) {
  const datos = e.parameter; // Obtiene los parámetros del formulario
  return guardarPedido(datos);
}

// Función para obtener estadísticas de pedidos
function obtenerEstadisticas(filtro) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const datos = hoja.getDataRange().getValues();
  const hoy = new Date();
  let pendientes = 0, entregados = 0;
  let totalIncidencias = 0, dcIncidencias = 0, montajeIncidencias = 0;

  datos.forEach((fila, index) => {
    if (index === 0) return; // Saltar encabezados
    const estado = fila[4]; // Columna del estado
    const incidencia = fila[5]; // Columna de incidencia (Sí/No)
    const tipoIncidencia = fila[6]; // Columna del tipo de incidencia
    const fechaEntrega = new Date(fila[1]);

    if (filtro === 'semana') {
      const diferenciaDias = (hoy - fechaEntrega) / (1000 * 60 * 60 * 24);
      if (diferenciaDias > 7) return;
    } else if (filtro === 'mes') {
      if (hoy.getMonth() !== fechaEntrega.getMonth() || hoy.getFullYear() !== fechaEntrega.getFullYear()) return;
    }

    // Contar estados
    if (estado === 'Pendiente') pendientes++;
    else if (estado === 'Entregado') entregados++;

    // Contar incidencias
    if (incidencia === 'Sí') {
      totalIncidencias++;
      if (tipoIncidencia === 'DC') dcIncidencias++;
      else if (tipoIncidencia === 'MONTAJE') montajeIncidencias++;
    }
  });

  const total = pendientes + entregados;
  return { pendientes, entregados, total, totalIncidencias, dcIncidencias, montajeIncidencias };
}

// Función para obtener resumen de estadísticas
function obtenerResumen(filtro) {
  const estadisticas = obtenerEstadisticas(filtro);
  return estadisticas;
}

// Función para buscar un pedido por número de pedido
function buscarPedido(numeroPedido) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const datos = hoja.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][1] === numeroPedido) { // Columna 2: Número de Pedido
      return {
        cliente: datos[i][0], // Nueva columna para el cliente
        pedido: datos[i][1],
        fechaPedido: datos[i][2],
        fechaEntrega: datos[i][3],
        incidencia: datos[i][5],
        tipoIncidencia: datos[i][6]
      };
    }
  }

  return null; // No se encontró el pedido
}

// Función para actualizar un pedido existente
function actualizarPedido(data) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const datos = hoja.getDataRange().getValues();

  for (let i = 1; i < datos.length; i++) {
    if (datos[i][1] === data.pedido) { // Columna 1: Número de Pedido
      hoja.getRange(i + 1, 1, 1, datos[i].length).setValues([[
        data.cliente || "",
        data.pedido || "",
        data.fechaPedido || "",
        data.fechaEntrega || "",
        data.diasDiferencia || "",
        data.estado || "",
        data.incidencia || "",
        data.tipoIncidencia || ""
      ]]);
      return 'Pedido actualizado correctamente.';
    }
  }

  throw new Error('No se encontró el pedido para actualizar.');
}



