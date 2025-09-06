/**
 * Funciones utilitarias comunes
 */

/**
 * Detecta índices de columnas basado en los encabezados
 */
function detectarIndicesColumnas(headers, columnConfig) {
  const indices = {};

  columnConfig.forEach((config, index) => {
    const foundIndex = headers.indexOf(config.name);
    indices[config.key] = foundIndex !== -1 ? foundIndex : index;
  });

  console.log(`Detectados índices: ${JSON.stringify(indices)}`);
  return indices;
}

/**
 * Extrae ID simplificado de ubicación completa
 */
function extraerIDSimple(idCompleto) {
  if (!idCompleto) return null;

  idCompleto = idCompleto.toString().trim();

  // Buscar patrones como L03-12 o similares
  const match = idCompleto.match(/^([A-Z][0-9]+(?:-[0-9]+)?)/);
  if (match) {
    return match[1];
  }

  // Si contiene guiones, tomar los primeros dos segmentos
  if (idCompleto.includes("-")) {
    const parts = idCompleto.split("-");
    if (parts.length >= 2) {
      return parts[0] + "-" + parts[1];
    }
  }

  return idCompleto;
}

/**
 * Genera colores para gradiente de ocupación
 */
function generarColorOcupacion(porcentaje, sinDatos = false) {
  const colors = getColors().OCUPACION;

  if (sinDatos) {
    return colors.VACIO;
  }

  porcentaje = Math.max(0, Math.min(100, porcentaje));

  if (porcentaje <= 25) return colors.BAJO;
  if (porcentaje <= 50) return colors.MEDIO;
  if (porcentaje <= 75) return colors.ALTO;
  return colors.COMPLETO;
}

/**
 * Genera colores para gradiente de vida útil
 */
function generarColorVidaUtil(porcentaje) {
  const colors = getColors().VENCIMIENTO;

  porcentaje = Math.max(0, Math.min(100, porcentaje));

  if (porcentaje >= 75) return colors.ALTO;
  if (porcentaje >= 50) return colors.MEDIO;
  if (porcentaje >= 25) return colors.BAJO;
  return colors.CRITICO;
}

/**
 * Formatea fecha para display
 */
function formatearFecha(fecha, formato = "dd/MM/yyyy") {
  if (!fecha || isNaN(fecha.getTime())) return "";

  return Utilities.formatDate(fecha, Session.getScriptTimeZone(), formato);
}

/**
 * Calcula días entre fechas
 */
function calcularDiasRestantes(fechaCaducidad, fechaActual = new Date()) {
  if (!fechaCaducidad || isNaN(fechaCaducidad.getTime())) return 0;

  const millisecondsPerDay = 1000 * 60 * 60 * 24;
  return Math.max(
    0,
    Math.floor((fechaCaducidad - fechaActual) / millisecondsPerDay)
  );
}

/**
 * Valida que los datos tengan el formato esperado
 */
function validarDatos(data, tipoData = "inventario") {
  if (!data || !Array.isArray(data) || data.length === 0) {
    throw new Error(
      `Los datos de ${tipoData} están vacíos o no tienen el formato esperado`
    );
  }

  if (data.length < 2) {
    throw new Error(
      `Los datos de ${tipoData} deben tener al menos encabezados y una fila de datos`
    );
  }

  return true;
}

/**
 * Limpia propiedades específicas del script
 */
function limpiarPropiedades(propiedades = []) {
  const scriptProperties = PropertiesService.getScriptProperties();

  const defaultProps = [
    "ultimoSVG_ID",
    "file_name",
    "ultimoSVG_URL",
    "processingCancelled",
    "isProcessingActive",
  ];

  const propsToDelete = propiedades.length > 0 ? propiedades : defaultProps;

  propsToDelete.forEach((prop) => {
    try {
      scriptProperties.deleteProperty(prop);
      console.log(`Propiedad ${prop} eliminada`);
    } catch (error) {
      console.log(`Error al eliminar propiedad ${prop}: ${error}`);
    }
  });
}

/**
 * Obtiene configuración de fuente de datos
 */
function obtenerConfiguracionFuenteDatos() {
  const scriptProperties = PropertiesService.getScriptProperties();
  return (
    scriptProperties.getProperty("fuenteDatosSeleccionada") || "blueyonder"
  );
}

/**
 * Actualiza configuración de fuente de datos
 */
function actualizarConfiguracionFuenteDatos(fuenteDatos) {
  try {
    PropertiesService.getScriptProperties().setProperty(
      "fuenteDatosSeleccionada",
      fuenteDatos
    );

    return {
      status: "success",
      message: "Fuente de datos actualizada",
    };
  } catch (error) {
    return {
      status: "error",
      message: "Error: " + error.toString(),
    };
  }
}

/**
 * Maneja errores de forma consistente
 */
function manejarError(error, contexto = "operación") {
  const mensaje = `Error en ${contexto}: ${error.toString()}`;
  console.error(mensaje);

  return {
    status: "error",
    message: mensaje,
  };
}

/**
 * Crea respuesta exitosa estandarizada
 */
function crearRespuestaExitosa(mensaje, datos = null) {
  const respuesta = {
    status: "success",
    message: mensaje,
  };

  if (datos) {
    respuesta.data = datos;
  }

  return respuesta;
}

/**
 * Valida archivo Excel
 */
function validarArchivoExcel(archivo) {
  if (!archivo) {
    throw new Error("No se proporcionó archivo");
  }

  const tiposValidos = [
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "application/vnd.ms-excel",
    MimeType.MICROSOFT_EXCEL,
    MimeType.MICROSOFT_EXCEL_LEGACY,
  ];

  const tipoArchivo = archivo.getContentType ? archivo.getContentType() : null;

  if (tipoArchivo && !tiposValidos.includes(tipoArchivo)) {
    console.log(`Tipo de archivo no válido: ${tipoArchivo}`);
    // No lanzar error, intentar procesar de todas formas
  }

  return true;
}

/**
 * Obtiene timestamp actual formateado
 */
function obtenerTimestampFormateado() {
  const fecha = new Date();
  const dia = fecha.getDate().toString().padStart(2, "0");
  const mes = (fecha.getMonth() + 1).toString().padStart(2, "0");
  const anio = fecha.getFullYear();
  const hora = fecha.getHours().toString().padStart(2, "0");
  const minutos = fecha.getMinutes().toString().padStart(2, "0");

  return `${dia}_${mes}_${anio}_${hora}_${minutos}`;
}
