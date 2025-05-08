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

  Logger.log("Datos recibidos:", data);
  
  const fechaEntrega = data.fechaEntrega || "";
  const fechaRuta = data.fechaRuta || "";

  const fechaEntregaObj = new Date(fechaEntrega);
  const fechaRutaObj = new Date(fechaRuta);

  Logger.log("Fecha Entrega:", fechaEntregaObj);
  Logger.log("Fecha Ruta:", fechaRutaObj);

  if (isNaN(fechaEntregaObj.getTime())) {
    Logger.log("La fecha de entrega no es válida.");
    return;
  }

  let estado = "Pendiente"; // Si no hay fecha de ruta, el estado es Pendiente.
  let diferenciaDias = "";
  let diferenciaHoras = "";

  if (fechaRuta) {
    const diferencia = fechaRutaObj - fechaEntregaObj; // Diferencia en milisegundos
    diferenciaHoras = diferencia / (1000 * 60 * 60); // Diferencia en horas
    diferenciaDias = Math.floor(diferencia / (1000 * 60 * 60 * 24)); // Diferencia en días sin decimales

    if (diferenciaHoras >= 15 && diferenciaHoras <= 20) {
      estado = "Entregado";
    } else if (diferenciaHoras > 20) {
      estado = "En Ruta";
    }
  }

  const incidenciaTexto = data.incidencia ? "Sí" : "No";
  const tipoIncidenciaTexto = data.incidencia && data.tipoIncidencia ? data.tipoIncidencia : "NO";

  hoja.appendRow([
    data.pedido || "",
    formatearFecha(fechaEntregaObj),
    fechaRuta ? formatearFecha(fechaRutaObj) : "", // Solo agrega la fecha de ruta si existe
    diferenciaDias || "", // Diferencia en días sin decimales
    estado, // Estado en "Pendiente", "En Ruta", o "Entregado"
    incidenciaTexto,
    tipoIncidenciaTexto // Si no hay incidencia, será "NO"
  ]);

  Logger.log("Pedido guardado correctamente.");
  return HtmlService.createHtmlOutput("¡Pedido guardado correctamente!").setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

// Función para manejar el POST que recibe los datos y los guarda en la hoja de Google Sheets
function doPost(e) {
  const datos = e.parameter; // Obtiene los parámetros del formulario
  return guardarPedido(datos);
}

