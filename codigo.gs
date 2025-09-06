/**
 * Constantes globales MODELO
 */
const SVG_BASE = "mapa_bodega.svg";
const FOLDER_ID = "1nrBTYyK24t0Yet6buZlzaiYY1NxJ9s3E";
const INVENTARIO_SHEET_ID = "1wPr43IgOOAW-vdv5sT7n3PrbOCbiF2lyWUlyvbu_iSU";
const SHEET_INVENTARIO = "InventarioBodega";
const SHEET_UBICACIONES = "Ubicaciones";
const SHEET_CAJAS_PALLET = "CajasxPallet";
const SVG_NAME = "mapa_bodega.svg";
const SUBCARPETA_EXCEL_ID = "1TpbKk1azJ0Aey-TIvj3N1Wu8IbEhNT5F";
const AREAS_CONSIDERADAS = [
  "A",
  "B",
  "C",
  "D",
  "E",
  "F",
  "G",
  "H",
  "I",
  "J",
  "Q",
  "R",
];
const PASILLO_RACK = ["A01"]; // No se cuenta en ubicaciones con menos de X pallets

/**
 * Constantes globales NPR
 */
// const SVG_BASE = "VNA SEMI_layout.svg";
// const SVG_BASE = "NPR_CD.svg";
// const FOLDER_ID = "11EkkJ6G1plZEv8asYje5N5B2_XPyTeYa";
// const INVENTARIO_SHEET_ID = "1mozdeFrNIjUxOYxuafWDbqUdI2JW2HmlhxJt2xVOK5Y";
// const SHEET_INVENTARIO = "InventarioBodega";
// const SHEET_UBICACIONES = "Ubicaciones";
// const SVG_NAME = "VNA SEMI_layout.svg";
// const SHEET_CAJAS_PALLET = "CajasxPallet";
// const SUBCARPETA_EXCEL_ID = "1J4502AyItDtmL-yf2o9y2moP4_ajvNpy";
// const AREAS_CONSIDERADAS = ["VNA SEMI AUTOMATICO", "AREA ALMACENAMIENTO"];

function detectarIndicesColumnasInventarioWEP(headers) {
  const indices = {
    ubicacion:
      headers.indexOf("Localidad") !== -1 ? headers.indexOf("Localidad") : 0,
    sku: headers.indexOf("SKU") !== -1 ? headers.indexOf("SKU") : 1,
    descripcion:
      headers.indexOf("Producto") !== -1 ? headers.indexOf("Producto") : 2,
    lpn: headers.indexOf("LPN") !== -1 ? headers.indexOf("LPN") : 3,
    cajas: headers.indexOf("Cantidad") !== -1 ? headers.indexOf("Cantidad") : 4,
    caducidad:
      headers.indexOf("Caducidad") !== -1 ? headers.indexOf("Caducidad") : 12,
  };
  console.log("Detectados índices Inventario WEP:", JSON.stringify(indices));
  return indices;
}

// Nombres columnas hoja ubicaciones WEP
function detectarIndicesColumnasUbicacionesWEP(headers) {
  const indices = {
    ubicacion:
      headers.indexOf("Localidad") !== -1 ? headers.indexOf("Localidad") : 0,
    area: headers.indexOf("Área") !== -1 ? headers.indexOf("Área") : 1,
    capacidad_maxima:
      headers.indexOf("Cap. Máxima") !== -1
        ? headers.indexOf("Cap. Máxima")
        : 3,
  };
  console.log("Detectados índices Ubicaciones WEP:", JSON.stringify(indices));
  return indices;
}

// Nombres columnas cuadratura inventario BlueYonder
function detectarIndicesColumnasInventario(headers) {
  const indices = {
    area: headers.indexOf("Área") !== -1 ? headers.indexOf("Área") : 1,
    zonaTrabajo:
      headers.indexOf("Zona de trabajo") !== -1
        ? headers.indexOf("Zona de trabajo")
        : 3,
    ubicacion:
      headers.indexOf("Ubicación") !== -1 ? headers.indexOf("Ubicación") : 4,
    lpn: headers.indexOf("LPN") !== -1 ? headers.indexOf("LPN") : 6,
    sku: headers.indexOf("Artículo") !== -1 ? headers.indexOf("Artículo") : 8,
    descripcion:
      headers.indexOf("Descripcion") !== -1
        ? headers.indexOf("Descripcion")
        : 10,
    cajas:
      headers.indexOf("Cantidad") !== -1 ? headers.indexOf("Cantidad") : 11,
    fifo:
      headers.indexOf("Fecha FIFO") !== -1 ? headers.indexOf("Fecha FIFO") : 14,
    fabricacion:
      headers.indexOf("Fecha de fabricacion") !== -1
        ? headers.indexOf("Fecha de fabricacion")
        : 15,
    caducidad:
      headers.indexOf("Fecha de caducidad") !== -1
        ? headers.indexOf("Fecha de caducidad")
        : 16,
    recepcion:
      headers.indexOf("Fecha de recepcion") !== -1
        ? headers.indexOf("Fecha de recepcion")
        : 17,
    perfilAntiguedad:
      headers.indexOf("Nombre de perfil de antiguedad") !== -1
        ? headers.indexOf("Nombre de perfil de antiguedad")
        : 19,
    diasHastaCaducidad:
      headers.indexOf("Dias hasta su caducidad") !== -1
        ? headers.indexOf("Dias hasta su caducidad")
        : 20,
    ultimoMov:
      headers.indexOf("Ultima fecha de movimiento") !== -1
        ? headers.indexOf("Ultima fecha de movimiento")
        : 21,
  };
  console.log("Detectados índices BlueYonder:", JSON.stringify(indices));
  return indices;
}

// Nombres columnas hoja ubicaciones BlueYonder
function detectarIndicesColumnasUbicaciones(headers) {
  const indices = {
    ubicacion:
      headers.indexOf("Ubicación") !== -1 ? headers.indexOf("Ubicación") : 1,
    area: headers.indexOf("Área") !== -1 ? headers.indexOf("Área") : 2,
  };
  console.log(
    "Detectados índices Ubicacioens BlueYonder:",
    JSON.stringify(indices)
  );
  return indices;
}

function detectarIndicesColumnasCajasxPallet(headers) {
  const indices = {
    sku: headers.indexOf("Artículo") !== -1 ? headers.indexOf("Artículo") : 0,
    cajas_x_pallet:
      headers.indexOf("Cajas X Pallets") !== -1
        ? headers.indexOf("Cajas X Pallets")
        : 1,
    vida_util:
      headers.indexOf("SKU_VIDA_UTIL") !== -1
        ? headers.indexOf("SKU_VIDA_UTIL")
        : 2,
  };
  console.log("Detectados índices CajasxPallet:", JSON.stringify(indices));
  return indices;
}

function obtenerUltimoArchivoExcel(carpetaId) {
  try {
    const carpeta = DriveApp.getFolderById(carpetaId);
    console.log(
      "Nombre de la carpeta: " +
        carpeta.getName() +
        ", URL: " +
        carpeta.getUrl()
    );

    let archivoMasReciente = null;
    let fechaMasReciente = new Date(0);

    // Buscar solo archivos que coincidan con nuestro patrón de nombres
    const allFiles = carpeta.getFiles();
    while (allFiles.hasNext()) {
      const archivo = allFiles.next();
      const nombre = archivo.getName();

      if (nombre.startsWith("Inventario_")) {
        const fechaModificacion = archivo.getLastUpdated();
        console.log(
          "Encontrado archivo de inventario: " +
            nombre +
            ", Fecha: " +
            fechaModificacion
        );

        if (fechaModificacion > fechaMasReciente) {
          fechaMasReciente = fechaModificacion;
          archivoMasReciente = archivo;
        }
      }
    }

    if (archivoMasReciente) {
      console.log(
        "Archivo Excel más reciente encontrado: " + archivoMasReciente.getName()
      );
      return archivoMasReciente;
    } else {
      console.log("No se encontró ningún archivo de inventario en la carpeta");
      return null;
    }
  } catch (error) {
    console.error(
      "Error al obtener el último archivo Excel: " + error.toString()
    );
    return null;
  }
}

function actualizarSVG(tipoMapa = "ocupacion") {
  limpiarPropiedadesEspecificas();
  try {
    if (typeof PropertiesService !== "undefined") {
      const docProps = PropertiesService.getDocumentProperties();
      if (docProps.getProperty("processingCancelled") === "true") {
        docProps.deleteProperty("processingCancelled");
        return {
          status: "cancelled",
          message: "Procesamiento cancelado por el usuario",
        };
      }
      // Establecer bandera de procesamiento
      docProps.setProperty("isProcessingActive", "true");
    }

    var documentProperties = PropertiesService.getDocumentProperties();
    fuenteDatos =
      documentProperties.getProperty("fuenteDatosSeleccionada") || "blueyonder";
    console.log("fuente datos", fuenteDatos);

    const ultimoArchivo = obtenerUltimoArchivoExcel(SUBCARPETA_EXCEL_ID);

    let data;
    if (!ultimoArchivo) {
      console.log("No hay ultimo archivo");
      // Si no hay archivo, usar la hoja de cálculo por defecto
      var sheet =
        SpreadsheetApp.openById(INVENTARIO_SHEET_ID).getSheetByName(
          SHEET_INVENTARIO
        );
      data = sheet.getDataRange().getValues();
    } else {
      // Convertir el archivo Excel a datos para procesar
      const blob = ultimoArchivo.getBlob();
      const datos = procesarArchivoExcelADatos(blob);

      if (!datos || !datos.inventario) {
        throw new Error("No se pudieron extraer datos del archivo Excel");
      }

      data = datos.inventario;
    }

    // Verificar que data tiene al menos una fila
    if (!data || !Array.isArray(data) || data.length === 0) {
      throw new Error("Los datos están vacíos o no tienen el formato esperado");
    }
    console.log("DATA:", data);

    // Obtengo ubicaciones desde hoja en drive
    var sheet_ubicaciones =
      SpreadsheetApp.openById(INVENTARIO_SHEET_ID).getSheetByName(
        SHEET_UBICACIONES
      );
    var data_ubicaciones = sheet_ubicaciones.getDataRange().getValues();

    // Verificar que data_ubicaciones tiene al menos una fila
    if (
      !data_ubicaciones ||
      !Array.isArray(data_ubicaciones) ||
      data_ubicaciones.length === 0
    ) {
      throw new Error(
        "Los datos de ubicaciones están vacíos o no tienen el formato esperado"
      );
    }

    var columnIndicesUbicaciones =
      fuenteDatos === "wep"
        ? detectarIndicesColumnasUbicacionesWEP(data_ubicaciones[0])
        : detectarIndicesColumnasUbicaciones(data_ubicaciones[0]);

    var sheet_ubicaciones =
      SpreadsheetApp.openById(INVENTARIO_SHEET_ID).getSheetByName(
        SHEET_UBICACIONES
      );
    var data_ubicaciones = sheet_ubicaciones.getDataRange().getValues();

    var sheetCajasxPallet =
      SpreadsheetApp.openById(INVENTARIO_SHEET_ID).getSheetByName(
        SHEET_CAJAS_PALLET
      );
    var data_cajas_x_pallet = sheetCajasxPallet.getDataRange().getValues();
    var columnIndicesCajasxPallet = detectarIndicesColumnasCajasxPallet(
      data_cajas_x_pallet[0]
    );
    // console.log("Indices Cajasxpallet:", columnIndicesCajasxPallet);

    // Cargar el SVG desde Drive
    // var svgFile = DriveApp.getFilesByName(SVG_NAME).next();
    // var svgContent = svgFile.getBlob().getDataAsString();

    // Procesar datos
    var ubicaciones_locations = procesarUbicaciones(
      data_ubicaciones,
      columnIndicesUbicaciones
    );
    // console.log("ubicaciones_locations:", ubicaciones_locations);

    // Usar los detectores de índices según la fuente de datos
    var columnIndicesInventario =
      fuenteDatos === "wep"
        ? detectarIndicesColumnasInventarioWEP(data[0])
        : detectarIndicesColumnasInventario(data[0]);

    var resultado = procesarUbicacionesMultinivel(
      ubicaciones_locations,
      data,
      data_cajas_x_pallet,
      columnIndicesInventario,
      columnIndicesCajasxPallet
    );
    console.log("ubicaciones multinivel listo");

    var ubicaciones = resultado.ubicaciones;
    var vencimientos = resultado.vencimientos;

    // var resultado = generarMapaColoresYDatos(ubicaciones, svgContent, "ocupacion", true);
    // console.log("Resuktado ocupacion mapa ocupacion listo");

    var datosUbicaciones = prepararDatosUbicaciones(ubicaciones);

    var timestamp = new Date().getTime();
    var filenameDatos = "datos_ubicaciones_" + timestamp + ".json";
    var saveResultDatos = guardarJSONEnDrive(
      datosUbicaciones,
      filenameDatos,
      FOLDER_ID
    );

    // var filenameOcupacion = "mapa_ocupacion_" + timestamp + ".svg";
    // console.log("filename", filenameOcupacion);
    // var saveResultOcupacion = guardarSVGEnDrive(FOLDER_ID, resultado.mapaOcupacion.svgContent, filenameOcupacion);
    // // console.log("save result ocupacion", saveResultOcupacion);

    // // Guardar mapa de vencimiento
    // var filenameVencimiento = "mapa_vencimiento_" + timestamp + ".svg";
    // var saveResultVencimiento = guardarSVGEnDrive(FOLDER_ID, resultado.mapaVencimiento.svgContent, filenameVencimiento);
    // console.log("Guardar mapa vencimeiento");

    // var filenameVencimiento = "mapa_vencimiento_" + timestamp + ".svg";
    // var saveResultVencimiento = guardarSVGEnDrive(FOLDER_ID, resultado.mapaVencimiento.svgContent, filenameVencimiento);
    // // console.log("guardar mapa vencimeiento", saveResultVencimiento);

    // // Guardar propiedades
    // var documentProperties = PropertiesService.getDocumentProperties();
    // documentProperties.setProperty('mapa_ocupacion_ID', saveResultOcupacion.id);
    // documentProperties.setProperty('mapa_ocupacion_URL', saveResultOcupacion.url);
    // documentProperties.setProperty('mapa_vencimiento_ID', saveResultVencimiento.id);
    // documentProperties.setProperty('mapa_vencimiento_URL', saveResultVencimiento.url);
    // console.log("guaradar prop");

    // Guardar mapas SKU
    var mapaSKU = generarMapaSKU(ubicaciones);
    var mapaSKUFileId = guardarJSONEnDrive(mapaSKU, "mapaSKU.json", FOLDER_ID);

    documentProperties.setProperty("datos_ubicaciones_ID", saveResultDatos);
    documentProperties.setProperty("mapaSKU_FileId", mapaSKUFileId);

    var vencimientosFileId = guardarJSONEnDrive(
      vencimientos,
      "datosVencimientos.json",
      FOLDER_ID
    );
    documentProperties.setProperty("vencimientosFileId", vencimientosFileId);

    if (typeof PropertiesService !== "undefined") {
      const docProps = PropertiesService.getDocumentProperties();
      docProps.deleteProperty("isProcessingActive");
    }

    return {
      status: "success",
      message: "Datos procesados correctamente.",
    };
  } catch (error) {
    Logger.log("Error: " + error.toString());
    return {
      status: "error",
      message: "Error: " + error.toString(),
    };
  }
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action === "update") {
    const result = actualizarSVG();
    return ContentService.createTextOutput(
      JSON.stringify({
        status: result.status,
        message: result.message,
      })
    ).setMimeType(ContentService.MimeType.JSON);
  }

  var svgFile = DriveApp.getFilesByName(SVG_BASE).next();
  var svgBaseContent = svgFile.getBlob().getDataAsString();

  // Obtener ID de los mapas
  var documentProperties = PropertiesService.getDocumentProperties();

  var mapaSKUFileId = documentProperties.getProperty("mapaSKU_FileId");
  var fuenteDatos =
    documentProperties.getProperty("fuenteDatosSeleccionada") || "blueyonder";

  var jsonSKU = obtenerMapaSKU(mapaSKUFileId);
  console.log("json vencimiento");

  var showHTML = e && e.parameter && e.parameter.html !== "false";

  if (showHTML) {
    // Devolver página HTML con SVG base
    var template = HtmlService.createTemplateFromFile("index");

    // Asignar variables
    template.TITULO = "Mapa de Bodega";
    template.FECHA_ACTUALIZACION = new Date().toLocaleString();
    template.SVG_OCUPACION = svgBaseContent;
    template.SVG_VENCIMIENTO = svgBaseContent;
    template.mapaSKU = JSON.stringify(jsonSKU || {});
    template.FUENTE_DATOS = fuenteDatos || "blueyonder";

    // Evaluar la plantilla
    var htmlOutput = template
      .evaluate()
      .setTitle("Mapa de Bodega")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    return htmlOutput;
  } else {
    // Para compatibilidad, devolver solo el SVG base
    return devolverSoloSVG(svgBaseContent);
  }
}

function devolverSoloSVG(svgContent) {
  if (svgContent.indexOf("<?xml") === -1) {
    svgContent =
      '<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n' + svgContent;
  }

  return ContentService.createTextOutput(svgContent).setMimeType(
    ContentService.MimeType.SVG
  );
}

function obtenerMapaSKU(mapaSKUFileId) {
  if (!mapaSKUFileId) {
    Logger.log("No mapaSKU file ID found");
    return {};
  }

  try {
    // Recuperar archivo
    var file = DriveApp.getFileById(mapaSKUFileId);
    var jsonContent = file.getBlob().getDataAsString();
    return JSON.parse(jsonContent);
  } catch (error) {
    Logger.log("Error al parsear el contenido de mapaSKU: " + error.toString());
    return {};
  }
}

function obtenerDatosVencimientos() {
  try {
    // Obtener ID del archivo
    var documentProperties = PropertiesService.getDocumentProperties();
    var vencimientosFileId =
      documentProperties.getProperty("vencimientosFileId");

    if (!vencimientosFileId) {
      Logger.log("No se encontró el archivo de vencimientos");
      return {};
    }

    // Recuperar datos
    var file = DriveApp.getFileById(vencimientosFileId);
    var jsonContent = file.getBlob().getDataAsString();

    try {
      var datosVencimientos = JSON.parse(jsonContent);
      Logger.log(
        "Datos de vencimientos cargados: " +
          Object.keys(datosVencimientos).length +
          " fechas"
      );
      return datosVencimientos;
    } catch (parseError) {
      Logger.log(
        "Error al parsear JSON de vencimientos: " + parseError.toString()
      );
      return {};
    }
  } catch (error) {
    Logger.log("Error al obtener datos de vencimientos: " + error.toString());
    return {};
  }
}

/**
 * Incluye contenido de archivos HTML
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Se llama en scriptsSlidePanel
function procesarArchivoExcelSeguro(datos) {
  try {
    console.log("Procesando archivo con método seguro");

    if (!datos || !datos.base64data || !datos.fileName) {
      return {
        status: "error",
        message: "Datos incompletos para procesamiento",
      };
    }

    // Obtener timestamp personalizado si existe
    const customTimestamp = datos.useCustomDate ? datos.timestamp : null;

    // Intentar decodificar el base64 de manera más robusta
    let decodedData;
    try {
      // Intentar decodificar el base64 directamente
      decodedData = Utilities.base64Decode(
        datos.base64data,
        Utilities.Charset.UTF_8
      );
    } catch (decodeError) {
      console.error("Error en la decodificación inicial:", decodeError);

      // Intentar con otro charset si falla
      try {
        decodedData = Utilities.base64Decode(datos.base64data);
      } catch (e) {
        throw new Error("No se pudo decodificar el archivo: " + e.message);
      }
    }

    // Verificar que tenemos datos después de la decodificación
    if (!decodedData || decodedData.length === 0) {
      throw new Error("La decodificación resultó en datos vacíos");
    }

    // Crear el blob con los datos decodificados
    const blob = Utilities.newBlob(
      decodedData,
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      datos.fileName
    );

    console.log(
      "Archivo decodificado correctamente, pasando a procesarArchivoExcelADatos"
    );

    // NO guardar el archivo aquí, dejar que procesarArchivoExcelADatos lo haga

    // Procesar el archivo usando la función existente, pasando el timestamp personalizado
    const resultado = procesarArchivoExcelADatos(blob, customTimestamp, false);

    return {
      status: "success",
      message: "Archivo procesado correctamente",
      processed: resultado.inventario ? resultado.inventario.length : 0,
    };
  } catch (error) {
    console.error("Error procesando archivo:", error);
    return {
      status: "error",
      message: "Error: " + error.toString(),
    };
  }
}

// Se llama en scripts al cambiar selector de fuente de datos
function actualizarFuenteDatos(fuenteDatos) {
  try {
    // Guardar la preferencia del usuario
    PropertiesService.getDocumentProperties().setProperty(
      "fuenteDatosSeleccionada",
      fuenteDatos
    );

    return {
      status: "success",
      message: "Fuente de datos actualizada",
    };
  } catch (error) {
    Logger.log("Error al actualizar fuente de datos: " + error.toString());
    return {
      status: "error",
      message: "Error: " + error.toString(),
    };
  }
}

// Se llama en scripts
function obtenerFuenteDatosSeleccionada() {
  var documentProperties = PropertiesService.getDocumentProperties();
  return (
    documentProperties.getProperty("fuenteDatosSeleccionada") || "blueyonder"
  );
}

// Se llama en scripts
function obtenerDatosUbicaciones() {
  try {
    // Obtener ID del archivo
    var documentProperties = PropertiesService.getDocumentProperties();
    var datosUbicacionesId = documentProperties.getProperty(
      "datos_ubicaciones_ID"
    );

    // Si no existe, procesar los datos bajo demanda
    if (!datosUbicacionesId) {
      console.log("No hay datos de ubicaciones, procesando bajo demanda");

      // Crear datos para el SVG base (datos vacíos)
      // Aquí se procesarían los datos igual que en actualizarSVG, pero sin guardar archivos

      // Como fallback, retornar un objeto vacío
      return {
        ocupacion: {},
        vencimiento: {},
      };
    }

    // Recuperar datos
    var file = DriveApp.getFileById(datosUbicacionesId);
    var jsonContent = file.getBlob().getDataAsString();

    try {
      var datosUbicaciones = JSON.parse(jsonContent);
      console.log("Datos de ubicaciones cargados correctamente");
      return datosUbicaciones;
    } catch (parseError) {
      console.log(
        "Error al parsear JSON de ubicaciones: " + parseError.toString()
      );
      return {
        ocupacion: {},
        vencimiento: {},
      };
    }
  } catch (error) {
    console.log("Error al obtener datos de ubicaciones: " + error.toString());
    return {
      ocupacion: {},
      vencimiento: {},
    };
  }
}

// Añadir esta función al final del archivo
function cancelarProcesamiento() {
  try {
    const docProps = PropertiesService.getDocumentProperties();

    // Establecer bandera de cancelación
    docProps.setProperty("processingCancelled", "true");

    // Verificar si hay procesamiento activo
    const isActive = docProps.getProperty("isProcessingActive") === "true";

    return {
      status: "success",
      message: isActive
        ? "Solicitud de cancelación enviada"
        : "No hay procesamiento activo",
      isActive: isActive,
    };
  } catch (error) {
    return {
      status: "error",
      message: "Error al cancelar: " + error.toString(),
    };
  }
}

/**
 * Obtiene la lista de archivos Excel en la carpeta de Drive
 */
function obtenerArchivosExcel() {
  try {
    const carpeta = DriveApp.getFolderById(SUBCARPETA_EXCEL_ID);
    const archivos = [];
    const archivosIds = new Set();

    // Tipos MIME para archivos Excel
    const tiposExcel = [
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", // .xlsx
      MimeType.MICROSOFT_EXCEL, // .xls
      MimeType.MICROSOFT_EXCEL_LEGACY, // .xls viejo
      "application/vnd.ms-excel", // Alternativo para .xls
    ];

    // Buscar archivos por cada tipo MIME
    tiposExcel.forEach((tipo) => {
      try {
        const fileIterator = carpeta.getFilesByType(tipo);

        while (fileIterator.hasNext()) {
          const file = fileIterator.next();
          const fileId = file.getId();

          if (!archivosIds.has(fileId)) {
            archivosIds.add(fileId);
            archivos.push({
              id: fileId,
              nombre: file.getName(),
              fecha: file.getLastUpdated().getTime(),
            });
          }
        }
      } catch (e) {
        console.log("Error al buscar tipo " + tipo + ": " + e.toString());
      }
    });

    // Ordenar por fecha (más reciente primero)
    archivos.sort((a, b) => b.fecha - a.fecha);

    return archivos;
  } catch (error) {
    console.error("Error al obtener archivos Excel:", error);
    return [];
  }
}

/**
 * Procesa un archivo Excel directamente desde Drive
 */
function procesarArchivoExcelDrive(fileId, customTimestamp, useExistingFile) {
  try {
    const file = DriveApp.getFileById(fileId);
    if (!file) {
      return {
        status: "error",
        message: "No se encontró el archivo en Drive",
      };
    }

    // Si useExistingFile es true, no crear un archivo nuevo
    // Procesamiento directo del archivo existente
    if (useExistingFile) {
      // Obtener un blob temporal para procesamiento
      const blob = file.getBlob();

      // En lugar de crear un nuevo archivo, actualizar la referencia
      // al último archivo Excel procesado
      if (SUBCARPETA_EXCEL_ID) {
        // Opcional: Guardar referencia en propiedades del documento
        PropertiesService.getDocumentProperties().setProperty(
          "ultimoArchivoExcelId",
          fileId
        );
      }

      // Usar función existente para procesar, pasando el timestamp
      const resultado = procesarArchivoExcelADatos(blob, customTimestamp, true);

      if (resultado.inventario) {
        return {
          status: "success",
          message: "Archivo existente procesado correctamente",
          processed: resultado.inventario.length,
        };
      } else {
        return {
          status: "error",
          message: "No se pudieron extraer datos del archivo existente",
        };
      }
    } else {
      // Código original para procesar y crear una copia del archivo
      // Este camino se seguiría cuando useExistingFile es false
      const blob = file.getBlob();
      const resultado = procesarArchivoExcelADatos(
        blob,
        customTimestamp,
        false
      );

      if (resultado.inventario) {
        return {
          status: "success",
          message: "Archivo procesado correctamente",
          processed: resultado.inventario.length,
        };
      } else {
        return {
          status: "error",
          message: "No se pudieron extraer datos del archivo",
        };
      }
    }
  } catch (error) {
    console.error("Error al procesar archivo de Drive:", error);
    return {
      status: "error",
      message: "Error: " + error.toString(),
    };
  }
}
