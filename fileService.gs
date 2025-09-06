/**
 * Servicio para operaciones con archivos en Google Drive
 */

/**
 * Obtiene el último archivo Excel de la carpeta
 */
function obtenerUltimoArchivoExcel(carpetaId = null) {
  try {
    // Primero intentar obtener el archivo por ID guardado
    const propiedades = gestionarPropiedadesDocumento();
    const ultimoArchivoId = propiedades.obtener("ultimoArchivoExcelId");
    
    if (ultimoArchivoId) {
      try {
        const archivo = DriveApp.getFileById(ultimoArchivoId);
        console.log(`Usando archivo guardado: ${archivo.getName()}`);
        return archivo;
      } catch (e) {
        console.log("No se pudo obtener archivo por ID guardado, buscando en carpeta...");
      }
    }
    
    // Si no hay ID guardado o falló, buscar en la carpeta
    const config = getConfig();
    const carpetaTarget = carpetaId || config.SUBCARPETA_EXCEL_ID;
    const carpeta = DriveApp.getFolderById(carpetaTarget);

    console.log(`Buscando archivos en carpeta: ${carpeta.getName()}`);

    let archivoMasReciente = null;
    let fechaMasReciente = new Date(0);

    const allFiles = carpeta.getFiles();
    while (allFiles.hasNext()) {
      const archivo = allFiles.next();
      const nombre = archivo.getName();

      if (nombre.startsWith("Inventario_") && (nombre.endsWith(".xlsx") || nombre.endsWith(".xls"))) {
        const fechaModificacion = archivo.getLastUpdated();

        if (fechaModificacion > fechaMasReciente) {
          fechaMasReciente = fechaModificacion;
          archivoMasReciente = archivo;
        }
      }
    }

    if (archivoMasReciente) {
      console.log(`Archivo más reciente: ${archivoMasReciente.getName()}`);
      // Guardar referencia para uso futuro
      propiedades.establecer("ultimoArchivoExcelId", archivoMasReciente.getId());
      return archivoMasReciente;
    }

    console.log("No se encontraron archivos de inventario");
    return null;
  } catch (error) {
    console.error(`Error al obtener archivo Excel: ${error}`);
    return null;
  }
}

/**
 * Verifica si un archivo es de tipo Excel
 */
function esArchivoExcel(archivo) {
  try {
    const tiposExcel = [
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "application/vnd.ms-excel",
      MimeType.MICROSOFT_EXCEL,
      MimeType.MICROSOFT_EXCEL_LEGACY,
    ];

    // Verificar si el archivo tiene el método getContentType
    if (!archivo || typeof archivo.getContentType !== 'function') {
      // Si no tiene getContentType, intentar verificar por el nombre
      if (archivo && typeof archivo.getName === 'function') {
        const nombre = archivo.getName();
        return nombre && (nombre.endsWith('.xlsx') || nombre.endsWith('.xls'));
      }
      return false;
    }

    const tipoArchivo = archivo.getContentType();
    return tiposExcel.includes(tipoArchivo);
  } catch (error) {
    console.error(`Error al verificar tipo de archivo: ${error}`);
    return false;
  }
}

/**
 * Obtiene lista de archivos Excel en la carpeta
 */
function obtenerListaArchivosExcel(carpetaId = null) {
  try {
    const config = getConfig();
    const carpetaTarget = carpetaId || config.SUBCARPETA_EXCEL_ID;
    const carpeta = DriveApp.getFolderById(carpetaTarget);

    const archivos = [];
    const archivosIds = new Set();

    const tiposExcel = [
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      MimeType.MICROSOFT_EXCEL,
      MimeType.MICROSOFT_EXCEL_LEGACY,
      "application/vnd.ms-excel",
    ];

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
              tamaño: file.getSize(),
            });
          }
        }
      } catch (e) {
        console.log(`Error al buscar tipo ${tipo}: ${e}`);
      }
    });

    // Ordenar por fecha (más reciente primero)
    archivos.sort((a, b) => b.fecha - a.fecha);

    return archivos;
  } catch (error) {
    console.error(`Error al obtener lista de archivos: ${error}`);
    return [];
  }
}

/**
 * Guarda datos JSON en Drive
 */
function guardarJSONEnDrive(datos, nombreArchivo, carpetaId = null) {
  try {
    const config = getConfig();
    const carpetaTarget = carpetaId || config.FOLDER_ID;
    const carpeta = DriveApp.getFolderById(carpetaTarget);

    const jsonString = JSON.stringify(datos, null, 2);
    const blob = Utilities.newBlob(
      jsonString,
      "application/json",
      nombreArchivo
    );

    // Eliminar archivo anterior si existe
    try {
      const archivosExistentes = carpeta.getFilesByName(nombreArchivo);
      while (archivosExistentes.hasNext()) {
        archivosExistentes.next().setTrashed(true);
      }
    } catch (e) {
      console.log(`No se pudo eliminar archivo anterior: ${e}`);
    }

    const archivo = carpeta.createFile(blob);
    console.log(`Archivo JSON guardado: ${nombreArchivo}`);

    return archivo.getId();
  } catch (error) {
    console.error(`Error al guardar JSON: ${error}`);
    throw error;
  }
}

/**
 * Carga datos JSON desde Drive
 */
function cargarJSONDesdeDrive(fileId) {
  try {
    if (!fileId) {
      console.log("No se proporcionó ID de archivo");
      return null;
    }

    const archivo = DriveApp.getFileById(fileId);
    const contenido = archivo.getBlob().getDataAsString();

    return JSON.parse(contenido);
  } catch (error) {
    console.error(`Error al cargar JSON: ${error}`);
    return null;
  }
}

/**
 * Guarda archivo SVG en Drive
 */
function guardarSVGEnDrive(carpetaId, contenidoSVG, nombreArchivo) {
  try {
    const carpeta = DriveApp.getFolderById(carpetaId);

    // Asegurar que el SVG tenga declaración XML
    let svgContent = contenidoSVG;
    if (svgContent.indexOf("<?xml") === -1) {
      svgContent =
        '<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n' + svgContent;
    }

    const blob = Utilities.newBlob(svgContent, "image/svg+xml", nombreArchivo);
    const archivo = carpeta.createFile(blob);

    console.log(`Archivo SVG guardado: ${nombreArchivo}`);

    return {
      id: archivo.getId(),
      url: archivo.getUrl(),
      nombre: nombreArchivo,
    };
  } catch (error) {
    console.error(`Error al guardar SVG: ${error}`);
    throw error;
  }
}

/**
 * Limpia archivos temporales de la carpeta
 */
function limpiarArchivosTemporales(carpetaId = null) {
  try {
    const config = getConfig();
    const carpetaTarget = carpetaId || config.SUBCARPETA_EXCEL_ID;
    const carpeta = DriveApp.getFolderById(carpetaTarget);

    const archivos = carpeta.getFiles();
    const archivosAEliminar = [];

    while (archivos.hasNext()) {
      const archivo = archivos.next();
      const nombreArchivo = archivo.getName();

      if (
        nombreArchivo.startsWith("temp_") ||
        nombreArchivo.includes("(Convertido)") ||
        nombreArchivo.includes("(Procesado)")
      ) {
        archivosAEliminar.push(archivo);
      }
    }

    // Eliminar archivos en lote
    archivosAEliminar.forEach((archivo) => {
      try {
        archivo.setTrashed(true);
        console.log(`Archivo temporal eliminado: ${archivo.getName()}`);
      } catch (error) {
        console.error(
          `Error al eliminar archivo ${archivo.getName()}: ${error}`
        );
      }
    });

    return archivosAEliminar.length;
  } catch (error) {
    console.error(`Error al limpiar archivos temporales: ${error}`);
    return 0;
  }
}

/**
 * Obtiene SVG base desde Drive
 */
function obtenerSVGBase() {
  try {
    const config = getConfig();
    const svgFiles = DriveApp.getFilesByName(config.SVG_BASE);

    if (!svgFiles.hasNext()) {
      throw new Error(`No se encontró el archivo SVG base: ${config.SVG_BASE}`);
    }

    const svgFile = svgFiles.next();
    return svgFile.getBlob().getDataAsString();
  } catch (error) {
    console.error(`Error al obtener SVG base: ${error}`);
    throw error;
  }
}

/**
 * Guarda archivo JSON de movimientos en Drive
 */
function guardarMovimientosEnDrive(datosMovimientos, timestamp = null) {
  try {
    const config = getConfig();
    const carpeta = DriveApp.getFolderById(config.SUBCARPETA_EXCEL_ID);
    const fechaActual = timestamp ? new Date(timestamp) : new Date();
    const nombreArchivo = `Movimientos_${formatearFecha(fechaActual, 'yyyy_MM_dd_HH_mm')}.json`;
    
    console.log('Guardando datos de movimientos:', {
      totalUbicaciones: Object.keys(datosMovimientos.movimientos || {}).length,
      estructura: Object.keys(datosMovimientos)
    });
    
    // Validar estructura de datos
    if (!datosMovimientos.movimientos) {
      console.warn('Los datos de movimientos no tienen la propiedad "movimientos"');
      datosMovimientos = {
        movimientos: datosMovimientos,
        estadisticas: { procesadoEn: new Date().toISOString() }
      };
    }
    
    // Crear blob con los datos
    const blob = Utilities.newBlob(
      JSON.stringify(datosMovimientos, null, 2),
      'application/json',
      nombreArchivo
    );
    
    // Limpiar archivos de movimientos anteriores (mantener solo el más reciente)
    limpiarArchivosMovimientosAnteriores(carpeta);
    
    // Crear nuevo archivo
    const archivo = carpeta.createFile(blob);
    console.log(`Archivo de movimientos guardado: ${nombreArchivo} con ${Object.keys(datosMovimientos.movimientos).length} ubicaciones`);
    
    // Guardar ID en propiedades del documento
    const propiedades = gestionarPropiedadesDocumento();
    propiedades.establecer('ultimoArchivoMovimientosId', archivo.getId());
    propiedades.establecer('ultimoArchivoMovimientosNombre', nombreArchivo);
    propiedades.establecer('fechaUltimoProcesamientoMovimientos', new Date().toISOString());
    
    // Log para verificar que se guardó correctamente
    console.log(`Propiedades actualizadas: ID=${archivo.getId()}, Nombre=${nombreArchivo}`);
    
    return {
      id: archivo.getId(),
      nombre: nombreArchivo,
      url: archivo.getUrl(),
      ubicacionesGuardadas: Object.keys(datosMovimientos.movimientos).length
    };
  } catch (error) {
    console.error(`Error al guardar movimientos: ${error}`);
    throw error;
  }
}

/**
 * Limpia archivos de movimientos anteriores para mantener solo el más reciente
 */
function limpiarArchivosMovimientosAnteriores(carpeta) {
  try {
    const archivosMovimientos = [];
    const archivos = carpeta.getFiles();
    
    while (archivos.hasNext()) {
      const archivo = archivos.next();
      const nombre = archivo.getName();
      
      if (nombre.startsWith('Movimientos_') && nombre.endsWith('.json')) {
        archivosMovimientos.push({
          archivo: archivo,
          nombre: nombre,
          fecha: archivo.getDateCreated()
        });
      }
    }
    
    // Ordenar por fecha y mantener solo los 2 más recientes
    archivosMovimientos.sort((a, b) => b.fecha - a.fecha);
    
    if (archivosMovimientos.length > 2) {
      for (let i = 2; i < archivosMovimientos.length; i++) {
        console.log(`Eliminando archivo de movimientos anterior: ${archivosMovimientos[i].nombre}`);
        archivosMovimientos[i].archivo.setTrashed(true);
      }
    }
  } catch (error) {
    console.warn(`Error limpiando archivos anteriores: ${error}`);
  }
}

/**
 * Carga archivo JSON de movimientos desde Drive
 */
function cargarMovimientosDesdeDrive() {
  try {
    const propiedades = gestionarPropiedadesDocumento();
    const archivoId = propiedades.obtener('ultimoArchivoMovimientosId');
    
    if (!archivoId) {
      console.log('No se encontró archivo de movimientos previo');
      return null;
    }
    
    console.log(`Intentando cargar movimientos desde archivo ID: ${archivoId}`);
    
    const archivo = DriveApp.getFileById(archivoId);
    const contenido = archivo.getBlob().getDataAsString();
    const datos = JSON.parse(contenido);
    
    const ubicacionesCount = datos.movimientos ? Object.keys(datos.movimientos).length : 0;
    console.log(`Movimientos cargados desde: ${archivo.getName()} con ${ubicacionesCount} ubicaciones`);
    
    // Debug: mostrar primeras ubicaciones
    if (datos.movimientos && ubicacionesCount > 0) {
      const primerasUbicaciones = Object.keys(datos.movimientos).slice(0, 3);
      console.log('Primeras ubicaciones cargadas:', primerasUbicaciones);
    }
    
    return datos;
    
  } catch (error) {
    console.error(`Error al cargar movimientos: ${error}`);
    console.error(`ID de archivo: ${propiedades.obtener('ultimoArchivoMovimientosId')}`);
    return null;
  }
}

/**
 * Convierte archivo Excel a Google Sheets
 */
function convertirExcelASheets(archivoExcel) {
  try {
    const config = getConfig();
    const carpeta = DriveApp.getFolderById(config.SUBCARPETA_EXCEL_ID);
    
    // Crear una copia del archivo Excel
    const nombreTemp = `temp_${Date.now()}_${archivoExcel.getName()}`;
    const blob = archivoExcel.getBlob();
    
    // Crear archivo temporal con el blob del Excel
    const archivoTemp = carpeta.createFile(blob);
    archivoTemp.setName(nombreTemp);
    
    // Obtener ID del archivo
    const fileId = archivoTemp.getId();
    
    // Usar Drive API para convertir a Google Sheets
    const resource = {
      title: nombreTemp.replace('.xlsx', '').replace('.xls', ''),
      mimeType: MimeType.GOOGLE_SHEETS,
      parents: [{id: config.SUBCARPETA_EXCEL_ID}]
    };
    
    try {
      // Convertir usando Drive API
      const convertedFile = Drive.Files.copy(resource, fileId, {convert: true});
      
      // Eliminar archivo temporal Excel
      archivoTemp.setTrashed(true);
      
      // Retornar el archivo convertido como SpreadsheetApp
      return SpreadsheetApp.openById(convertedFile.id);
      
    } catch (apiError) {
      console.error('Error con Drive API, intentando método alternativo:', apiError);
      
      // Método alternativo: crear un nuevo spreadsheet y copiar datos
      const newSpreadsheet = SpreadsheetApp.create(nombreTemp);
      const newSpreadsheetFile = DriveApp.getFileById(newSpreadsheet.getId());
      
      // Mover a la carpeta correcta
      carpeta.addFile(newSpreadsheetFile);
      DriveApp.getRootFolder().removeFile(newSpreadsheetFile);
      
      // Eliminar archivo temporal
      archivoTemp.setTrashed(true);
      
      return newSpreadsheet;
    }
    
  } catch (error) {
    console.error(`Error al convertir Excel a Sheets: ${error}`);
    throw new Error(`No se pudo convertir el archivo Excel: ${error.message}`);
  }
}

/**
 * Gestiona propiedades del script
 * Usa ScriptProperties ya que no hay documento asociado
 */
function gestionarPropiedadesDocumento() {
  return {
    obtener: (clave) => {
      try {
        return PropertiesService.getScriptProperties().getProperty(clave);
      } catch (error) {
        console.error(`Error al obtener propiedad ${clave}: ${error}`);
        return null;
      }
    },

    establecer: (clave, valor) => {
      try {
        PropertiesService.getScriptProperties().setProperty(clave, valor);
        return true;
      } catch (error) {
        console.error(`Error al establecer propiedad ${clave}: ${error}`);
        return false;
      }
    },

    eliminar: (clave) => {
      try {
        PropertiesService.getScriptProperties().deleteProperty(clave);
        return true;
      } catch (error) {
        console.error(`Error al eliminar propiedad ${clave}: ${error}`);
        return false;
      }
    },
  };
}
