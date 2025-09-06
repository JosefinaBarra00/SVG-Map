/**
 * Aplicación principal - Controlador y puntos de entrada
 */

/**
 * Punto de entrada principal para la aplicación web
 */
function doGet(e) {
  try {
    // Manejar actualización via API
    if (e && e.parameter && e.parameter.action === "update") {
      const result = procesarActualizacion();
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    // Obtener datos necesarios para la página
    const datosApp = obtenerDatosAplicacion();
    
    // Verificar si debe devolver HTML o SVG
    const mostrarHTML = e && e.parameter && e.parameter.html !== "false";
    
    if (mostrarHTML) {
      return generarRespuestaHTML(datosApp);
    } else {
      return generarRespuestaSVG(datosApp.svgBase);
    }
  } catch (error) {
    console.error(`Error en doGet: ${error}`);
    return ContentService.createTextOutput(
      JSON.stringify(manejarError(error, "aplicación web"))
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Obtiene datos necesarios para la aplicación
 */
function obtenerDatosAplicacion() {
  const propiedades = gestionarPropiedadesDocumento();
  
  return {
    svgBase: obtenerSVGBase(),
    mapaSKU: obtenerMapaSKUDesdeArchivo(propiedades.obtener("mapaSKU_FileId")),
    fuenteDatos: propiedades.obtener("fuenteDatosSeleccionada") || "blueyonder",
    fechaActualizacion: new Date().toLocaleString(),
    movimientos: cargarMovimientosDesdeDrive()
  };
}

/**
 * Genera respuesta HTML para la aplicación
 */
function generarRespuestaHTML(datosApp) {
  const template = HtmlService.createTemplateFromFile("index");
  
  // Asignar variables a la plantilla
  template.TITULO = "Mapa de Bodega";
  template.FECHA_ACTUALIZACION = datosApp.fechaActualizacion;
  template.SVG_OCUPACION = datosApp.svgBase;
  template.SVG_VENCIMIENTO = datosApp.svgBase;
  template.mapaSKU = JSON.stringify(datosApp.mapaSKU || {});
  template.FUENTE_DATOS = datosApp.fuenteDatos;
  
  return template.evaluate()
    .setTitle("Mapa de Bodega")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Genera respuesta SVG simple
 */
function generarRespuestaSVG(svgContent) {
  let contenido = svgContent;
  
  if (contenido.indexOf("<?xml") === -1) {
    contenido = '<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n' + contenido;
  }
  
  return ContentService.createTextOutput(contenido)
    .setMimeType(ContentService.MimeType.SVG);
}

/**
 * Procesa actualización de datos
 */
function procesarActualizacion() {
  try {
    limpiarPropiedades();
    
    // Verificar si hay cancelación pendiente
    const propiedades = gestionarPropiedadesDocumento();
    if (propiedades.obtener("processingCancelled") === "true") {
      propiedades.eliminar("processingCancelled");
      return crearRespuestaExitosa("Procesamiento cancelado por el usuario");
    }
    
    // Establecer estado de procesamiento
    propiedades.establecer("isProcessingActive", "true");
    
    try {
      const resultado = ejecutarProcesamiento();
      return resultado;
    } finally {
      propiedades.eliminar("isProcessingActive");
    }
  } catch (error) {
    return manejarError(error, "actualización de datos");
  }
}

/**
 * Ejecuta el procesamiento principal de datos
 */
function ejecutarProcesamiento() {
  const fuenteDatos = obtenerConfiguracionFuenteDatos();
  console.log(`Procesando con fuente de datos: ${fuenteDatos}`);
  
  // Obtener datos de inventario del Excel
  const datosInventario = obtenerDatosInventario();
  
  // Extraer ubicaciones únicas del inventario
  const datosUbicaciones = extraerUbicacionesDeInventario(datosInventario);
  
  // Extraer datos de SKU del inventario (o usar array vacío si no hay)
  const datosSKU = extraerSKUsDeInventario(datosInventario);
  
  // Procesar todos los datos
  const resultado = procesarDatosCompletos(
    datosInventario,
    datosUbicaciones, 
    datosSKU,
    fuenteDatos
  );
  
  // Guardar resultados procesados
  guardarResultadosProcesados(resultado);
  
  return crearRespuestaExitosa("Datos procesados correctamente");
}

/**
 * Obtiene datos de inventario (desde archivo Excel)
 */
function obtenerDatosInventario() {
  const ultimoArchivo = obtenerUltimoArchivoExcel();
  
  if (!ultimoArchivo) {
    throw new Error("No se encontró ningún archivo Excel de inventario. Por favor, suba un archivo Excel para procesar.");
  }
  
  // Procesar archivo Excel
  const datosExcel = procesarArchivoExcel(ultimoArchivo);
  
  if (!datosExcel || !datosExcel.inventario) {
    throw new Error("No se pudieron extraer datos del archivo Excel");
  }
  
  return datosExcel.inventario;
}

/**
 * Extrae ubicaciones únicas del inventario
 */
function extraerUbicacionesDeInventario(datosInventario) {
  if (!datosInventario || datosInventario.length < 2) {
    return [["Ubicación", "Área", "Capacidad Máxima"]]; // Headers only
  }
  
  const ubicacionesMap = new Map();
  const headers = datosInventario[0];
  
  // Buscar índices de columnas manualmente
  const indices = {
    ubicacion: buscarIndiceColumna(headers, ["Ubicación", "Localidad", "Location"]),
    area: buscarIndiceColumna(headers, ["Área", "Area", "Zone"])
  };
  
  // Procesar filas (saltando headers)
  for (let i = 1; i < datosInventario.length; i++) {
    const fila = datosInventario[i];
    const ubicacion = fila[indices.ubicacion];
    const area = fila[indices.area] || "";
    
    if (ubicacion && !ubicacionesMap.has(ubicacion)) {
      ubicacionesMap.set(ubicacion, {
        ubicacion: ubicacion,
        area: area,
        capacidad_maxima: 1 // Capacidad por defecto
      });
    }
  }
  
  // Convertir a array formato tabla
  const resultado = [["Ubicación", "Área", "Capacidad Máxima"]];
  ubicacionesMap.forEach(ub => {
    resultado.push([ub.ubicacion, ub.area, ub.capacidad_maxima]);
  });
  
  console.log(`Extraídas ${ubicacionesMap.size} ubicaciones únicas del inventario`);
  return resultado;
}

/**
 * Extrae SKUs únicos del inventario
 */
function extraerSKUsDeInventario(datosInventario) {
  if (!datosInventario || datosInventario.length < 2) {
    return [["Artículo", "Cajas X Pallets", "SKU_VIDA_UTIL"]]; // Headers only
  }
  
  const skuMap = new Map();
  const headers = datosInventario[0];
  
  // Buscar índices de columnas manualmente
  const indices = {
    sku: buscarIndiceColumna(headers, ["Artículo", "SKU", "Item", "Producto"]),
    descripcion: buscarIndiceColumna(headers, ["Descripcion", "Description", "Desc"]),
    cajas: buscarIndiceColumna(headers, ["Cantidad", "Cajas", "Boxes", "Qty"])
  };
  
  // Procesar filas (saltando headers)
  for (let i = 1; i < datosInventario.length; i++) {
    const fila = datosInventario[i];
    const sku = fila[indices.sku];
    
    if (sku && !skuMap.has(sku)) {
      skuMap.set(sku, {
        sku: sku,
        cajas_x_pallet: 1, // Valor por defecto
        vida_util: 365 // Vida útil por defecto en días
      });
    }
  }
  
  // Convertir a array formato tabla
  const resultado = [["Artículo", "Cajas X Pallets", "SKU_VIDA_UTIL"]];
  skuMap.forEach(item => {
    resultado.push([item.sku, item.cajas_x_pallet, item.vida_util]);
  });
  
  console.log(`Extraídos ${skuMap.size} SKUs únicos del inventario`);
  return resultado;
}

/**
 * Busca el índice de una columna en los headers
 */
function buscarIndiceColumna(headers, posiblesNombres) {
  for (let i = 0; i < headers.length; i++) {
    const header = (headers[i] || "").toString().trim();
    for (const nombre of posiblesNombres) {
      if (header.toLowerCase().includes(nombre.toLowerCase())) {
        return i;
      }
    }
  }
  return -1; // No encontrado
}



/**
 * Procesa todos los datos de forma integrada
 */
function procesarDatosCompletos(datosInventario, datosUbicaciones, datosSKU, fuenteDatos) {
  // Validar datos
  validarDatos(datosInventario, "inventario");
  validarDatos(datosUbicaciones, "ubicaciones");
  validarDatos(datosSKU, "SKU");
  
  // Detectar índices de columnas
  const indicesInventario = detectarIndicesColumnas(
    datosInventario[0], 
    getColumnIndices(fuenteDatos, "inventario")
  );
  
  const indicesUbicaciones = detectarIndicesColumnas(
    datosUbicaciones[0],
    getColumnIndices(fuenteDatos, "ubicaciones")
  );
  
  const indicesSKU = detectarIndicesColumnas(
    datosSKU[0],
    getColumnIndices(fuenteDatos, "cajas_pallet")
  );
  
  // Procesar ubicaciones base
  const ubicacionesBase = procesarUbicacionesBase(datosUbicaciones, indicesUbicaciones);
  
  // Procesar datos multinivel
  const resultado = procesarUbicacionesMultinivel(
    ubicacionesBase,
    datosInventario,
    datosSKU,
    indicesInventario,
    indicesSKU
  );
  
  // Preparar datos para visualización
  const datosVisualizacion = prepararDatosVisualizacion(resultado.ubicaciones);
  
  // Generar mapa de SKU
  const mapaSKU = generarMapaSKU(resultado.ubicaciones);
  
  return {
    ubicaciones: resultado.ubicaciones,
    vencimientos: resultado.vencimientos,
    datosVisualizacion: datosVisualizacion,
    mapaSKU: mapaSKU
  };
}

/**
 * Guarda los resultados procesados en Drive
 */
function guardarResultadosProcesados(resultado) {
  const timestamp = new Date().getTime();
  const propiedades = gestionarPropiedadesDocumento();
  
  // Guardar datos de ubicaciones
  const idDatosUbicaciones = guardarJSONEnDrive(
    resultado.datosVisualizacion,
    `datos_ubicaciones_${timestamp}.json`
  );
  
  // Guardar mapa de SKU
  const idMapaSKU = guardarJSONEnDrive(
    resultado.mapaSKU,
    "mapaSKU.json"
  );
  
  // Guardar datos de vencimientos
  const idVencimientos = guardarJSONEnDrive(
    resultado.vencimientos,
    "datosVencimientos.json"
  );
  
  // Actualizar propiedades del documento
  propiedades.establecer("datos_ubicaciones_ID", idDatosUbicaciones);
  propiedades.establecer("mapaSKU_FileId", idMapaSKU);
  propiedades.establecer("vencimientosFileId", idVencimientos);
  
  console.log("Resultados guardados exitosamente");
}

/**
 * Incluye archivos HTML (utilidad para plantillas)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Procesa archivo Excel subido por el usuario
 */
function procesarArchivoSubido(datos) {
  try {
    console.log("Procesando archivo subido de forma segura");
    
    if (!datos || !datos.base64data || !datos.fileName) {
      return manejarError(new Error("Datos incompletos para procesamiento"), "subida de archivo");
    }
    
    const timestampPersonalizado = datos.useCustomDate ? datos.timestamp : null;
    
    // Decodificar archivo
    const archivoDecodificado = decodificarArchivoBase64(datos.base64data, datos.fileName);
    
    // Validar que sea un archivo Excel
    if (!datos.fileName.match(/\.(xlsx|xls)$/i)) {
      return manejarError(new Error("El archivo debe ser Excel (.xlsx o .xls)"), "validación de archivo");
    }
    
    // Procesar archivo
    const resultado = procesarArchivoExcel(archivoDecodificado, timestampPersonalizado);
    
    // Guardar referencia al último archivo procesado
    const propiedades = gestionarPropiedadesDocumento();
    if (resultado && resultado.inventario) {
      // El archivo se procesó correctamente, ahora ejecutar actualización completa
      console.log("Archivo Excel procesado, ejecutando actualización completa del mapa...");
      
      // Ejecutar el procesamiento completo que actualiza el mapa
      const resultadoActualizacion = procesarActualizacion();
      
      if (resultadoActualizacion.status === "success") {
        return crearRespuestaExitosa(
          "Archivo procesado y mapa actualizado correctamente",
          { 
            procesadas: resultado.inventario.length,
            mensaje: "El mapa de bodega ha sido actualizado con los nuevos datos"
          }
        );
      } else {
        return resultadoActualizacion;
      }
    }
    
    return crearRespuestaExitosa(
      "Archivo procesado correctamente",
      { procesadas: resultado.inventario ? resultado.inventario.length : 0 }
    );
  } catch (error) {
    return manejarError(error, "procesamiento de archivo subido");
  }
}

/**
 * Decodifica archivo base64 de forma robusta
 */
function decodificarArchivoBase64(base64data, fileName) {
  try {
    let datosDecodificados = Utilities.base64Decode(base64data, Utilities.Charset.UTF_8);
    
    if (!datosDecodificados || datosDecodificados.length === 0) {
      throw new Error("La decodificación resultó en datos vacíos");
    }
    
    return Utilities.newBlob(
      datosDecodificados,
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      fileName
    );
  } catch (error) {
    // Intentar con charset diferente
    try {
      const datosDecodificados = Utilities.base64Decode(base64data);
      return Utilities.newBlob(
        datosDecodificados,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        fileName
      );
    } catch (e) {
      throw new Error(`No se pudo decodificar el archivo: ${e.message}`);
    }
  }
}

/**
 * Valida que el archivo sea un Excel válido
 */
function validarArchivoExcel(archivo) {
  const tiposValidos = [
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "application/vnd.ms-excel",
    MimeType.MICROSOFT_EXCEL,
    MimeType.MICROSOFT_EXCEL_LEGACY
  ];
  
  const tipoArchivo = archivo.getMimeType();
  if (!tiposValidos.includes(tipoArchivo)) {
    throw new Error(`Tipo de archivo no válido: ${tipoArchivo}. Se requiere un archivo Excel (.xlsx o .xls)`);
  }
  
  return true;
}

/**
 * Procesa archivo Excel desde Drive
 */
function procesarArchivoDrive(fileId, customTimestamp, useExistingFile) {
  try {
    const archivo = DriveApp.getFileById(fileId);
    if (!archivo) {
      return manejarError(new Error("Archivo no encontrado en Drive"), "archivo de Drive");
    }
    
    validarArchivoExcel(archivo);
    
    const resultado = procesarArchivoExcel(archivo, customTimestamp, useExistingFile);
    
    if (useExistingFile) {
      // Actualizar referencia al último archivo procesado
      const propiedades = gestionarPropiedadesDocumento();
      propiedades.establecer("ultimoArchivoExcelId", fileId);
    }
    
    return crearRespuestaExitosa(
      "Archivo de Drive procesado correctamente",
      { procesadas: resultado.inventario ? resultado.inventario.length : 0 }
    );
  } catch (error) {
    return manejarError(error, "procesamiento de archivo de Drive");
  }
}

/**
 * Cancela procesamiento en curso
 */
function cancelarProcesamiento() {
  try {
    const propiedades = gestionarPropiedadesDocumento();
    propiedades.establecer("processingCancelled", "true");
    
    const isActive = propiedades.obtener("isProcessingActive") === "true";
    
    return crearRespuestaExitosa(
      isActive ? "Solicitud de cancelación enviada" : "No hay procesamiento activo",
      { isActive }
    );
  } catch (error) {
    return manejarError(error, "cancelación de procesamiento");
  }
}

/**
 * Obtiene mapa de SKU desde archivo
 */
function obtenerMapaSKUDesdeArchivo(mapaSKUFileId) {
  if (!mapaSKUFileId) {
    console.log("No hay ID de archivo para mapa de SKU");
    return {};
  }
  
  const datos = cargarJSONDesdeDrive(mapaSKUFileId);
  return datos || {};
}

/**
 * Obtiene mapa de SKU usando las propiedades del documento automáticamente
 */
function obtenerMapaSKUAutomatico() {
  try {
    const propiedades = gestionarPropiedadesDocumento();
    const mapaSKUFileId = propiedades.obtener("mapaSKU_FileId");
    
    if (!mapaSKUFileId) {
      console.log("No hay ID de archivo para mapa de SKU en propiedades");
      return {};
    }
    
    const datos = cargarJSONDesdeDrive(mapaSKUFileId);
    return datos || {};
  } catch (error) {
    console.error(`Error al obtener mapa de SKU: ${error}`);
    return {};
  }
}

/**
 * Procesa movimientos desde archivo de inventario (subido o existente)
 */
function procesarArchivoMovimientos(base64Data, filename, periodoInicio, periodoFin, tipoAnalisis) {
  try {
    console.log("Procesando movimientos desde inventario:", filename || "archivo existente");
    
    // Verificar cancelación
    const propiedades = gestionarPropiedadesDocumento();
    if (propiedades.obtener("processingCancelled") === "true") {
      propiedades.eliminar("processingCancelled");
      return crearRespuestaExitosa("Procesamiento cancelado por el usuario");
    }
    
    let archivoInventario;
    let esModoSubida = base64Data && filename;
    
    if (esModoSubida) {
      // Modo 1: Procesar archivo subido
      console.log("Procesando archivo subido:", filename);
      
      // Convertir base64 a blob
      const archivoDecodificado = decodificarArchivoBase64Inventario(base64Data, filename);
      
      // Validar que sea CSV
      if (!filename.toLowerCase().endsWith('.csv')) {
        return manejarError(new Error("El archivo debe ser CSV (.csv) para análisis de movimientos"), "validación de archivo");
      }
      
      // Guardar temporalmente para procesamiento
      archivoInventario = guardarArchivoInventarioTemporal(archivoDecodificado, filename);
      
    } else {
      // Modo 2: Usar archivo existente
      console.log("Buscando archivo de inventario existente...");
      archivoInventario = obtenerArchivoInventarioCSV();
      
      if (!archivoInventario) {
        return manejarError(
          new Error("No se encontró archivo de inventario. Por favor, suba un archivo CSV de inventario."), 
          "búsqueda de archivo"
        );
      }
    }
    
    console.log("Usando archivo de inventario:", archivoInventario.getName());
    
    // Procesar movimientos desde el archivo de inventario
    const resultadoMovimientos = procesarMovimientosDesdeInventario(archivoInventario, periodoInicio, periodoFin, tipoAnalisis);
    
    // Limpiar archivo temporal si se subió uno nuevo
    if (esModoSubida && archivoInventario.getId) {
      try {
        archivoInventario.setTrashed(true);
      } catch (e) {
        console.warn("No se pudo eliminar archivo temporal:", e);
      }
    }
    
    // Guardar resultado en Drive
    const archivoGuardado = guardarMovimientosEnDrive(resultadoMovimientos);
    
    // Verificar que se guardó correctamente
    const verificacion = verificarGuardadoMovimientos();
    
    return crearRespuestaExitosa(
      esModoSubida ? "Archivo de inventario procesado y movimientos analizados" : "Movimientos procesados desde inventario existente",
      {
        archivo: archivoGuardado,
        estadisticas: resultadoMovimientos.estadisticas,
        totalUbicaciones: Object.keys(resultadoMovimientos.movimientos).length,
        fuente: archivoInventario.getName(),
        modo: esModoSubida ? "archivo_subido" : "archivo_existente",
        verificacionGuardado: verificacion
      }
    );
  } catch (error) {
    return manejarError(error, "procesamiento de movimientos desde inventario");
  }
}

/**
 * Decodifica archivo base64 para inventario CSV
 */
function decodificarArchivoBase64Inventario(base64data, fileName) {
  try {
    let datosDecodificados = Utilities.base64Decode(base64data, Utilities.Charset.UTF_8);
    
    if (!datosDecodificados || datosDecodificados.length === 0) {
      throw new Error("La decodificación resultó en datos vacíos");
    }
    
    return Utilities.newBlob(
      datosDecodificados,
      "text/csv",
      fileName
    );
  } catch (error) {
    // Intentar con charset diferente
    try {
      const datosDecodificados = Utilities.base64Decode(base64data);
      return Utilities.newBlob(
        datosDecodificados,
        "text/csv",
        fileName
      );
    } catch (e) {
      throw new Error(`No se pudo decodificar el archivo CSV: ${e.message}`);
    }
  }
}

/**
 * Guarda archivo de inventario temporal en Drive
 */
function guardarArchivoInventarioTemporal(blob, nombreArchivo) {
  try {
    const config = getConfig();
    const carpeta = DriveApp.getFolderById(config.SUBCARPETA_EXCEL_ID);
    
    // Crear nombre temporal único
    const timestamp = Date.now();
    const nombreTemp = `inventario_temp_${timestamp}_${nombreArchivo}`;
    
    const archivoCreado = carpeta.createFile(blob.setName(nombreTemp));
    
    console.log(`Archivo inventario temporal guardado: ${nombreTemp} con ID: ${archivoCreado.getId()}`);
    
    return archivoCreado;
  } catch (error) {
    console.error(`Error guardando archivo temporal: ${error}`);
    throw new Error(`Error guardando archivo temporal: ${error.message}`);
  }
}

/**
 * Obtiene el archivo CSV de inventario desde la carpeta raíz del proyecto
 */
function obtenerArchivoInventarioCSV() {
  try {
    // Buscar archivo Inventario.csv en la carpeta actual
    const archivos = DriveApp.getFilesByName("Inventario.csv");
    
    if (archivos.hasNext()) {
      const archivo = archivos.next();
      console.log(`Archivo inventario encontrado: ${archivo.getName()}, ID: ${archivo.getId()}`);
      return archivo;
    }
    
    // Si no se encuentra, buscar en la carpeta de Excel como respaldo
    const config = getConfig();
    if (config.SUBCARPETA_EXCEL_ID) {
      const carpeta = DriveApp.getFolderById(config.SUBCARPETA_EXCEL_ID);
      const archivosEnCarpeta = carpeta.getFilesByName("Inventario.csv");
      
      if (archivosEnCarpeta.hasNext()) {
        const archivo = archivosEnCarpeta.next();
        console.log(`Archivo inventario encontrado en carpeta Excel: ${archivo.getName()}`);
        return archivo;
      }
    }
    
    console.warn("No se encontró archivo Inventario.csv");
    return null;
  } catch (error) {
    console.error(`Error buscando archivo de inventario: ${error}`);
    return null;
  }
}

/**
 * Procesa movimientos desde archivo de inventario CSV
 */
function procesarMovimientosDesdeInventario(archivoInventario, periodoInicio, periodoFin, tipoAnalisis) {
  try {
    console.log("Procesando movimientos desde inventario CSV...");
    
    // Leer contenido del archivo CSV
    const csvContent = archivoInventario.getBlob().getDataAsString('UTF-8');
    const data = procesarCSVInventario(csvContent);
    
    console.log(`Datos CSV procesados: ${data.length} filas`);
    
    // Detectar índices de columnas específicos para inventario
    const indices = detectarIndicesInventarioMovimientos(data[0]);
    console.log('Índices detectados para inventario:', JSON.stringify(indices));
    
    // Procesar movimientos con filtrado de fechas
    const movimientos = procesarDatosMovimientosInventario(data, indices, periodoInicio, periodoFin, tipoAnalisis);
    
    return movimientos;
    
  } catch (error) {
    console.error(`Error procesando movimientos desde inventario: ${error}`);
    throw error;
  }
}

/**
 * Procesa contenido CSV de inventario con manejo mejorado de multi-línea
 */
function procesarCSVInventario(csvContent) {
  try {
    console.log("Iniciando procesamiento de CSV de inventario...");
    
    // Separar líneas pero mantener información de saltos de línea
    const lines = csvContent.split(/\r?\n/);
    console.log(`Total líneas detectadas: ${lines.length}`);
    
    // Reconstruir filas respetando campos entre comillas multi-línea
    const rows = [];
    let current = '';
    let inQuote = false;
    
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      
      // Contar comillas dobles en la línea
      const quoteCount = (line.match(/"/g) || []).length;
      
      if (!inQuote) {
        current = line;
      } else {
        // Si estamos dentro de comillas, agregar la línea con un espacio
        current += ' ' + line;
      }
      
      // Cambiar estado si hay número impar de comillas
      inQuote ^= quoteCount % 2 !== 0;
      
      // Si no estamos en comillas, es una fila completa
      if (!inQuote && current.trim()) {
        rows.push(current);
        current = '';
      }
    }
    
    // Agregar la última fila si queda algo
    if (current.trim()) {
      rows.push(current);
    }
    
    console.log(`Filas reconstruidas: ${rows.length}`);
    
    // Procesar cada fila reconstruida
    const data = [];
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i].trim();
      if (row) {
        try {
          const parsedRow = parseCSVLineInventarioMejorado(row);
          if (parsedRow.length > 0) {
            data.push(parsedRow);
            
            // Log de las primeras filas para debug
            if (i < 3) {
              console.log(`Fila ${i} parseada:`, JSON.stringify(parsedRow.slice(0, 5)));
            }
          }
        } catch (parseError) {
          console.warn(`Error parseando fila ${i}: ${parseError.message}`);
          // Continuar con la siguiente fila en lugar de fallar completamente
        }
      }
    }
    
    console.log(`CSV de inventario procesado exitosamente: ${data.length} filas válidas`);
    return data;
    
  } catch (error) {
    console.error(`Error procesando CSV de inventario: ${error}`);
    throw new Error(`Error procesando archivo CSV de inventario: ${error.message}`);
  }
}

/**
 * Parsea línea CSV de inventario con manejo especial para campos multi-línea (versión original)
 */
function parseCSVLineInventario(line) {
  const result = [];
  let current = '';
  let inQuotes = false;
  
  for (let i = 0; i < line.length; i++) {
    const char = line[i];
    
    if (char === '"') {
      inQuotes = !inQuotes;
    } else if (char === ',' && !inQuotes) {
      result.push(current.trim());
      current = '';
    } else {
      current += char;
    }
  }
  
  // Agregar el último campo
  result.push(current.trim());
  
  // Limpiar comillas de los campos y normalizar
  return result.map(field => {
    let cleaned = field.replace(/^"|"$/g, '');
    // Reemplazar saltos de línea internos con espacios
    cleaned = cleaned.replace(/\r?\n/g, ' ').trim();
    return cleaned;
  });
}

/**
 * Parsea línea CSV mejorado con mejor manejo de comillas y espacios
 */
function parseCSVLineInventarioMejorado(line) {
  const result = [];
  let current = '';
  let inQuotes = false;
  let i = 0;
  
  // Limpiar la línea de espacios extra y saltos de línea
  const cleanLine = line.replace(/\r?\n/g, ' ').replace(/\s+/g, ' ').trim();
  
  while (i < cleanLine.length) {
    const char = cleanLine[i];
    
    if (char === '"') {
      // Manejar comillas dobles escapadas
      if (i + 1 < cleanLine.length && cleanLine[i + 1] === '"' && inQuotes) {
        current += '"';
        i += 2; // Saltar ambas comillas
        continue;
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === ',' && !inQuotes) {
      // Separador de campo encontrado
      result.push(limpiarCampoCSV(current));
      current = '';
      i++;
      continue;
    } else {
      current += char;
    }
    
    i++;
  }
  
  // Agregar el último campo
  result.push(limpiarCampoCSV(current));
  
  return result;
}

/**
 * Limpia un campo CSV individual
 */
function limpiarCampoCSV(field) {
  if (!field) return '';
  
  let cleaned = field.trim();
  
  // Remover comillas del inicio y final si están balanceadas
  if (cleaned.startsWith('"') && cleaned.endsWith('"')) {
    cleaned = cleaned.slice(1, -1);
  }
  
  // Limpiar espacios múltiples
  cleaned = cleaned.replace(/\s+/g, ' ').trim();
  
  return cleaned;
}

/**
 * Detecta índices de columnas para archivo de inventario CSV
 */
function detectarIndicesInventarioMovimientos(headers) {
  console.log('Headers de inventario detectados:', JSON.stringify(headers));
  
  // Mapeo específico para el formato de inventario CSV
  const columnMapping = {
    transactionDate: ['Fecha de transacción', 'Fecha de transaccion', 'fecha de transaccion'],
    user: ['Usuario', 'usuario'],
    activity: ['Actividad', 'actividad'],
    operation: ['Operación', 'Operacion', 'operacion'],
    item: ['Artículo', 'Articulo', 'articulo', 'Item'],
    quantity: ['Cantidad', 'cantidad'],
    moveUnit: ['Mover UM', 'mover um'],
    lpn: ['LPN'],
    subLpn: ['Sub-LPN', 'sub-lpn'],
    detailLpn: ['LPN de detalle', 'lpn de detalle'],
    fromLocation: ['Ubicación de origen', 'ubicacion de origen'],
    toLocation: ['Ubicación de destino', 'ubicacion de destino'],
    fromArea: ['Área de origen', 'area de origen'],
    toArea: ['Área de destino', 'area de destino'],
    movementRef: ['Referencia de movimiento', 'referencia de movimiento']
  };
  
  const indices = {};
  
  // Normalizar headers
  const headersNormalizados = headers.map(h => normalizarTexto(h || ''));
  
  for (const [key, possibleNames] of Object.entries(columnMapping)) {
    indices[key] = -1;
    
    for (let i = 0; i < headers.length; i++) {
      const headerOriginal = (headers[i] || '').toLowerCase().trim();
      const headerNormalizado = headersNormalizados[i];
      
      for (const nombre of possibleNames) {
        const nombreNormalizado = normalizarTexto(nombre);
        
        if (headerNormalizado.includes(nombreNormalizado) || 
            headerOriginal.includes(nombreNormalizado) ||
            headerOriginal.includes(nombre.toLowerCase())) {
          indices[key] = i;
          console.log(`Columna "${key}" encontrada en posición ${i}: "${headers[i]}"`);
          break;
        }
      }
      
      if (indices[key] !== -1) break;
    }
    
    if (indices[key] === -1) {
      console.warn(`Columna "${key}" no encontrada en headers de inventario`);
    }
  }
  
  console.log('Índices finales para inventario:', JSON.stringify(indices));
  return indices;
}

/**
 * Verifica que los movimientos se hayan guardado correctamente
 */
function verificarGuardadoMovimientos() {
  try {
    const propiedades = gestionarPropiedadesDocumento();
    const archivoId = propiedades.obtener('ultimoArchivoMovimientosId');
    const fechaGuardado = propiedades.obtener('fechaUltimoProcesamientoMovimientos');
    
    if (!archivoId) {
      return {
        status: "error",
        mensaje: "No se encontró ID de archivo de movimientos en propiedades"
      };
    }
    
    // Intentar cargar el archivo
    const datosMovimientos = cargarMovimientosDesdeDrive();
    if (!datosMovimientos) {
      return {
        status: "error",
        mensaje: "No se pudieron cargar los datos de movimientos guardados"
      };
    }
    
    const ubicacionesCount = datosMovimientos.movimientos ? Object.keys(datosMovimientos.movimientos).length : 0;
    
    return {
      status: "success",
      mensaje: "Movimientos guardados y verificados correctamente",
      detalles: {
        archivoId: archivoId,
        fechaGuardado: fechaGuardado,
        ubicacionesGuardadas: ubicacionesCount,
        tieneEstructuraCorrecta: !!datosMovimientos.movimientos
      }
    };
    
  } catch (error) {
    return {
      status: "error",
      mensaje: `Error verificando guardado: ${error.message}`
    };
  }
}

/**
 * Obtiene datos de movimientos procesados
 */
function obtenerDatosMovimientos() {
  try {
    const movimientos = cargarMovimientosDesdeDrive();
    
    if (!movimientos) {
      return crearRespuestaExitosa("No hay datos de movimientos", { movimientos: null });
    }
    
    return crearRespuestaExitosa("Datos de movimientos obtenidos", movimientos);
  } catch (error) {
    return manejarError(error, "obtención de datos de movimientos");
  }
}

/**
 * Funciones exportadas para uso en scripts HTML
 */

// Exportar funciones para uso en scripts
function actualizarSVG(tipoMapa = "ocupacion") {
  return procesarActualizacion();
}

function procesarArchivoExcelSeguro(datos) {
  return procesarArchivoSubido(datos);
}

function actualizarFuenteDatos(fuenteDatos) {
  return actualizarConfiguracionFuenteDatos(fuenteDatos);
}

function obtenerFuenteDatosSeleccionada() {
  return obtenerConfiguracionFuenteDatos();
}

function obtenerDatosUbicacionesParaUI() {
  try {
    const propiedades = gestionarPropiedadesDocumento();
    const datosUbicacionesId = propiedades.obtener("datos_ubicaciones_ID");
    
    // Datos base de ocupación y vencimiento
    let datos = { ocupacion: {}, vencimiento: {}, movimientos: {} };
    
    console.log("Obteniendo datos para UI...");
    
    if (datosUbicacionesId) {
      const datosGuardados = cargarJSONDesdeDrive(datosUbicacionesId);
      if (datosGuardados) {
        datos.ocupacion = datosGuardados.ocupacion || {};
        datos.vencimiento = datosGuardados.vencimiento || {};
        console.log(`Datos base cargados: ${Object.keys(datos.ocupacion).length} ubicaciones ocupación, ${Object.keys(datos.vencimiento).length} ubicaciones vencimiento`);
      }
    } else {
      console.log("No se encontró ID de datos de ubicaciones");
    }
    
    // Agregar datos de movimientos si existen
    const datosMovimientos = cargarMovimientosDesdeDrive();
    if (datosMovimientos && datosMovimientos.movimientos) {
      datos.movimientos = datosMovimientos.movimientos;
      console.log(`Datos de movimientos agregados: ${Object.keys(datos.movimientos).length} ubicaciones con movimientos`);
      
      // Debug: mostrar algunos datos de movimientos
      const primerasUbicaciones = Object.keys(datos.movimientos).slice(0, 3);
      primerasUbicaciones.forEach(ubicacion => {
        const mov = datos.movimientos[ubicacion];
        console.log(`Ubicación ${ubicacion}: entradas=${mov.entradas}, salidas=${mov.salidas}`);
      });
    } else {
      console.log("No se encontraron datos de movimientos o estructura incorrecta");
    }
    
    console.log(`Datos finales para UI: ocupacion=${Object.keys(datos.ocupacion).length}, vencimiento=${Object.keys(datos.vencimiento).length}, movimientos=${Object.keys(datos.movimientos).length}`);
    
    return datos;
  } catch (error) {
    console.error(`Error al obtener datos de ubicaciones: ${error}`);
    return { ocupacion: {}, vencimiento: {}, movimientos: {} };
  }
}

function obtenerArchivosExcel() {
  return obtenerListaArchivosExcel();
}

function procesarArchivoExcelDrive(fileId, customTimestamp, useExistingFile) {
  return procesarArchivoDrive(fileId, customTimestamp, useExistingFile);
}

function actualizarTipoMapa(tipoMapa) {
  try {
    const propiedades = gestionarPropiedadesDocumento();
    propiedades.establecer("tipoMapaSeleccionado", tipoMapa);
    
    return crearRespuestaExitosa("Tipo de mapa actualizado");
  } catch (error) {
    return manejarError(error, "actualización de tipo de mapa");
  }
}

function obtenerDatosVencimientos() {
  try {
    const propiedades = gestionarPropiedadesDocumento();
    const vencimientosFileId = propiedades.obtener("vencimientosFileId");
    
    if (!vencimientosFileId) {
      console.log("No se encontró archivo de vencimientos");
      return {};
    }
    
    const datos = cargarJSONDesdeDrive(vencimientosFileId);
    return datos || {};
  } catch (error) {
    console.error(`Error al obtener datos de vencimientos: ${error}`);
    return {};
  }
}

/**
 * Obtiene SVG base para el frontend
 */
function obtenerSVGBaseParaUI() {
  try {
    return obtenerSVGBase();
  } catch (error) {
    console.error(`Error al obtener SVG base: ${error}`);
    return "";
  }
}