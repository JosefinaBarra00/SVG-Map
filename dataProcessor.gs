/**
 * Procesador de datos optimizado y refactorizado
 */

/**
 * Procesa archivo de inventario (Excel o CSV)
 */
function procesarArchivoExcel(blob, customTimestamp = null, useExistingFile = false) {
  const startTime = Date.now();
  console.log(`Iniciando procesamiento de archivo de inventario: ${customTimestamp}`);
  
  try {
    if (useExistingFile) {
      return procesarArchivoExistente(blob);
    }
    
    const timestamp = customTimestamp || Date.now();
    const fileName = blob.getName() || '';
    const isCSV = fileName.toLowerCase().endsWith('.csv');
    
    let inventarioData;
    
    if (isCSV) {
      // Procesar CSV de inventario con transacciones
      console.log('Procesando archivo CSV de inventario...');
      inventarioData = procesarInventarioCSV(blob, timestamp);
    } else {
      // Procesar Excel tradicional
      console.log('Procesando archivo Excel de inventario...');
      const nombreArchivo = generarNombreArchivo(timestamp);
      const archivoTemporal = guardarArchivoTemporal(blob, nombreArchivo);
      inventarioData = procesarConCache(archivoTemporal, nombreArchivo);
    }
    
    const processingTime = Date.now() - startTime;
    console.log(`Archivo procesado exitosamente en ${processingTime}ms`);
    
    return validarYRetornarDatos(inventarioData);
    
  } catch (error) {
    console.error(`Error en procesamiento: ${error}`);
    throw error;
  }
}

/**
 * Genera nombre de archivo con timestamp
 */
function generarNombreArchivo(timestamp) {
  const fecha = new Date(timestamp);
  const componentes = {
    dia: fecha.getDate().toString().padStart(2, "0"),
    mes: (fecha.getMonth() + 1).toString().padStart(2, "0"),
    anio: fecha.getFullYear(),
    hora: fecha.getHours().toString().padStart(2, "0"),
    minutos: fecha.getMinutes().toString().padStart(2, "0")
  };
  
  return `Inventario_${componentes.dia}_${componentes.mes}_${componentes.anio}_${componentes.hora}_${componentes.minutos}.xlsx`;
}

/**
 * Guarda archivo temporal en Drive
 */
function guardarArchivoTemporal(blob, nombreArchivo) {
  const config = getConfig();
  const carpeta = DriveApp.getFolderById(config.SUBCARPETA_EXCEL_ID);
  
  // Limpiar archivos temporales antes de crear uno nuevo
  limpiarArchivosTemporales();
  
  const archivoCreado = carpeta.createFile(blob.setName(nombreArchivo));
  
  // Guardar ID del último archivo procesado
  const propiedades = gestionarPropiedadesDocumento();
  propiedades.establecer("ultimoArchivoExcelId", archivoCreado.getId());
  propiedades.establecer("ultimoArchivoExcelNombre", nombreArchivo);
  
  console.log(`Archivo guardado: ${nombreArchivo} con ID: ${archivoCreado.getId()}`);
  
  return archivoCreado;
}

/**
 * Procesa archivo existente sin crear copia
 */
function procesarArchivoExistente(archivo) {
  try {
    let driveFile;
    
    // Determinar si es un blob o un archivo de Drive
    if (archivo.getId) {
      // Es un archivo de Drive
      driveFile = archivo;
    } else if (archivo.getBlob) {
      // Es un archivo que puede proporcionar un blob
      driveFile = archivo;
    } else {
      throw new Error("Formato de archivo no reconocido");
    }
    
    // Procesar directamente el archivo Excel
    return procesarConCache(driveFile, driveFile.getName());
  } catch (error) {
    console.error(`Error al procesar archivo existente: ${error}`);
    throw error;
  }
}

/**
 * Procesa archivo con sistema de caché
 */
function procesarConCache(archivo, nombreArchivo) {
  const config = getProcessingConfig();
  const cacheKey = `processed_${nombreArchivo}_${archivo.getId()}`;
  
  let cache = null;
  let cachedData = null;
  
  // Intentar obtener cache de forma segura
  try {
    cache = CacheService.getDocumentCache();
    if (cache) {
      cachedData = cache.get(cacheKey);
      if (cachedData) {
        console.log(`Usando datos en caché para: ${nombreArchivo}`);
        return JSON.parse(cachedData);
      }
    } else {
      console.warn("CacheService.getDocumentCache() retornó null, continuando sin caché");
    }
  } catch (cacheError) {
    console.warn(`Error accediendo al cache: ${cacheError}, continuando sin caché`);
    cache = null;
  }
  
  // Procesar archivo
  console.log("Iniciando procesamiento del archivo...");
  
  // Procesar archivo Excel directamente
  const inventarioData = procesarExcelDirectamente(archivo);
  
  console.log(`Datos extraídos exitosamente. Filas: ${inventarioData.length}`);
  
  // Guardar en caché si está disponible
  const datosParaCache = { inventario: inventarioData };
  if (cache) {
    try {
      cache.put(cacheKey, JSON.stringify(datosParaCache), config.CACHE_DURATION);
      console.log(`Datos guardados en caché con clave: ${cacheKey}`);
    } catch (cacheError) {
      console.warn(`Error guardando en caché: ${cacheError}`);
    }
  }
  
  return datosParaCache;
}

/**
 * Procesa archivo Excel directamente usando DriveApp
 */
function procesarExcelDirectamente(archivo) {
  try {
    console.log(`Procesando archivo Excel: ${archivo.getName()}`);
    
    // Convertir Excel a Google Sheets
    const hojaConvertida = convertirExcelASheets(archivo);
    
    // Esperar un momento para que la conversión se complete
    const config = getProcessingConfig();
    Utilities.sleep(config.SHEET_WAIT);
    
    // Obtener la primera hoja (o la hoja de inventario)
    let hoja = hojaConvertida.getSheets()[0];
    
    // Buscar hoja específica si existe
    const nombresHojaInventario = ['Inventario', 'Inventory', 'Data', 'Datos', 'Sheet1', 'Hoja1'];
    for (const nombre of nombresHojaInventario) {
      const hojaTemp = hojaConvertida.getSheetByName(nombre);
      if (hojaTemp) {
        hoja = hojaTemp;
        console.log(`Usando hoja: ${nombre}`);
        break;
      }
    }
    
    // Obtener todos los datos
    const datos = hoja.getDataRange().getValues();
    
    if (!datos || datos.length === 0) {
      throw new Error("La hoja no contiene datos");
    }
    
    console.log(`Datos extraídos: ${datos.length} filas, ${datos[0].length} columnas`);
    
    // Eliminar el archivo temporal de Google Sheets después de procesarlo
    try {
      const archivoTemporal = DriveApp.getFileById(hojaConvertida.getId());
      archivoTemporal.setTrashed(true);
      console.log("Archivo temporal de Sheets eliminado");
    } catch (deleteError) {
      console.warn(`No se pudo eliminar archivo temporal: ${deleteError}`);
    }
    
    return datos;
    
  } catch (error) {
    console.error(`Error al procesar Excel directamente: ${error}`);
    throw new Error(`Error procesando archivo Excel: ${error.message}`);
  }
}


/**
 * Valida y retorna datos procesados
 */
function validarYRetornarDatos(datosCache) {
  // Si datosCache ya tiene la estructura { inventario: [...] }, usar directamente
  const inventarioData = datosCache.inventario || datosCache;
  
  if (inventarioData && Array.isArray(inventarioData) && inventarioData.length > 1) {
    console.log("Primera fila:", JSON.stringify(inventarioData[0]));
    console.log("Segunda fila:", JSON.stringify(inventarioData[1]));
    return { inventario: inventarioData };
  } else {
    throw new Error("La hoja no contiene suficientes datos");
  }
}

/**
 * Procesa ubicaciones base
 */
function procesarUbicacionesBase(dataUbicaciones, indices) {
  const ubicaciones = {};
  
  for (let i = 1; i < dataUbicaciones.length; i++) {
    const location = dataUbicaciones[i][indices.ubicacion];
    if (!location) continue;
    
    const idRack = extraerIDSimple(location);
    const area = dataUbicaciones[i][indices.area] || "";
    
    if (!ubicaciones[idRack]) {
      ubicaciones[idRack] = {
        area: area,
        zona_trabajo: "",
        capacidad_maxima: dataUbicaciones[i][indices.capacidad_maxima] || 0,
        utilizado: 0,
        niveles: {}
      };
    }
    
    if (!ubicaciones[idRack].niveles[location]) {
      ubicaciones[idRack].niveles[location] = {
        utilizado: 0,
        skus: {}
      };
    }
    
    // Actualizar capacidad máxima contando niveles
    if (ubicaciones[idRack].capacidad_maxima === 0) {
      ubicaciones[idRack].capacidad_maxima = Object.keys(ubicaciones[idRack].niveles).length;
    }
  }
  
  return ubicaciones;
}

/**
 * Procesa ubicaciones multinivel optimizado
 */
function procesarUbicacionesMultinivel(ubicaciones, data, dataSKU, indicesInventario, indicesSKU) {
  const startTime = Date.now();
  const vencimientos = {};
  
  // Crear mapa de vida útil
  const skuVidaUtilMap = crearMapaVidaUtil(dataSKU, indicesSKU);
  const hoy = new Date();
  
  // Procesar en lotes
  procesarDatosEnLotes(data, ubicaciones, vencimientos, skuVidaUtilMap, hoy, indicesInventario);
  
  const processingTime = Date.now() - startTime;
  console.log(`Procesamiento multinivel completado en ${processingTime}ms`);
  
  return { ubicaciones, vencimientos };
}

/**
 * Crea mapa de vida útil de SKUs
 */
function crearMapaVidaUtil(dataSKU, indicesSKU) {
  const skuVidaUtilMap = {};
  
  if (dataSKU && dataSKU.length > 0) {
    for (let i = 1; i < dataSKU.length; i++) {
      if (dataSKU[i] && dataSKU[i][indicesSKU.sku]) {
        skuVidaUtilMap[dataSKU[i][indicesSKU.sku]] = Number(dataSKU[i][indicesSKU.vida_util]) || 0;
      }
    }
  }
  
  return skuVidaUtilMap;
}

/**
 * Procesa datos en lotes optimizados
 */
function procesarDatosEnLotes(data, ubicaciones, vencimientos, skuVidaUtilMap, hoy, indicesInventario) {
  const config = getProcessingConfig();
  const batchSize = config.BATCH_SIZE;
  const totalBatches = Math.ceil((data.length - 1) / batchSize);
  
  for (let i = 1; i < data.length; i += batchSize) {
    const endIdx = Math.min(i + batchSize, data.length);
    const batchNumber = Math.floor(i / batchSize) + 1;
    
    console.log(`Procesando lote ${batchNumber}/${totalBatches} (${endIdx - i} filas)`);
    
    procesarLoteOptimizado(
      data.slice(i, endIdx),
      ubicaciones,
      vencimientos,
      skuVidaUtilMap,
      hoy,
      indicesInventario
    );
  }
}

/**
 * Procesa lote de datos optimizado
 */
function procesarLoteOptimizado(loteData, ubicaciones, vencimientos, skuVidaUtilMap, hoy, indicesInventario) {
  const millisecondsPerDay = 1000 * 60 * 60 * 24;
  const timeZone = Session.getScriptTimeZone();
  
  for (let j = 0; j < loteData.length; j++) {
    const row = loteData[j];
    const location = row[indicesInventario.ubicacion];
    
    if (!location) continue;
    
    const idRack = extraerIDSimple(location);
    const cantidad = Number(row[indicesInventario.cajas]) > 0 ? 1 : 0;
    const sku = row[indicesInventario.sku] ? String(row[indicesInventario.sku]) : "";
    
    // Inicializar ubicación y nivel
    inicializarUbicacionYNivel(ubicaciones, idRack, location, row, indicesInventario, cantidad);
    
    // Procesar SKU si existe
    if (sku && cantidad > 0) {
      procesarSKUEnNivel(
        ubicaciones[idRack].niveles[location],
        sku,
        row,
        indicesInventario,
        cantidad,
        hoy,
        skuVidaUtilMap,
        millisecondsPerDay,
        timeZone
      );
    }
    
    // Procesar vencimientos
    procesarVencimiento(row, location, vencimientos, indicesInventario, timeZone);
  }
}

/**
 * Inicializa ubicación y nivel
 */
function inicializarUbicacionYNivel(ubicaciones, idRack, location, row, indicesInventario, cantidad) {
  if (!ubicaciones[idRack]) {
    ubicaciones[idRack] = {
      area: row[indicesInventario.area] || "",
      zona_trabajo: row[indicesInventario.zona_trabajo] || "",
      capacidad_maxima: 1,
      utilizado: 0,
      niveles: {}
    };
  }
  
  ubicaciones[idRack].utilizado += cantidad;
  
  if (!ubicaciones[idRack].niveles[location]) {
    ubicaciones[idRack].niveles[location] = {
      utilizado: 0,
      skus: {}
    };
  }
}

/**
 * Procesa SKU en nivel específico
 */
function procesarSKUEnNivel(nivel, sku, row, indicesInventario, cantidad, hoy, skuVidaUtilMap, millisecondsPerDay, timeZone) {
  if (!nivel.skus[sku]) {
    nivel.skus[sku] = {
      sku: sku,
      descripcion: row[indicesInventario.descripcion] ? String(row[indicesInventario.descripcion]) : "",
      cajas: 0,
      pallets: 0,
      lpns: [],
      fechas: []
    };
  }
  
  const skuObj = nivel.skus[sku];
  const cajas = Number(row[indicesInventario.cajas]) || 0;
  
  // Actualizar cantidades
  skuObj.cajas += cajas;
  skuObj.pallets += cantidad;
  
  // Agregar LPN si no existe
  const lpn = row[indicesInventario.lpn] ? String(row[indicesInventario.lpn]) : "";
  if (lpn && skuObj.lpns.indexOf(lpn) === -1) {
    skuObj.lpns.push(lpn);
  }
  
  // Procesar fechas y vida útil
  procesarFechasYVidaUtil(row, skuObj, hoy, skuVidaUtilMap, sku, indicesInventario, timeZone, millisecondsPerDay);
  
  nivel.utilizado += cantidad;
}

/**
 * Procesa fechas y vida útil
 */
function procesarFechasYVidaUtil(row, skuObj, hoy, skuVidaUtilMap, sku, indicesInventario, timeZone, millisecondsPerDay) {
  if (!row[indicesInventario.caducidad]) return;
  
  const fechaCaducidad = new Date(row[indicesInventario.caducidad]);
  if (!fechaCaducidad || isNaN(fechaCaducidad.getTime())) return;
  
  const fechaFormateada = formatearFecha(fechaCaducidad, "dd/MM/yyyy");
  
  if (fechaFormateada && skuObj.fechas.indexOf(fechaFormateada) === -1) {
    skuObj.fechas.push(fechaFormateada);
  }
  
  let diasRestantes, vidaUtilTotal;
  
  if (row[indicesInventario.diasHastaCaducidad]) {
    diasRestantes = row[indicesInventario.diasHastaCaducidad];
    vidaUtilTotal = row[indicesInventario.perfilAntiguedad];
  } else {
    diasRestantes = calcularDiasRestantes(fechaCaducidad, hoy);
    vidaUtilTotal = skuVidaUtilMap[sku] || 0;
  }
  
  if (vidaUtilTotal > 0) {
    skuObj.vida_util_dias = diasRestantes;
    skuObj.vida_util_porc = Math.min(100, Math.round((diasRestantes / vidaUtilTotal) * 100));
  }
}

/**
 * Procesa vencimientos
 */
function procesarVencimiento(row, location, vencimientos, indicesInventario, timeZone) {
  const fechaCaducidad = new Date(row[indicesInventario.caducidad]);
  if (!fechaCaducidad || isNaN(fechaCaducidad.getTime())) return;
  
  const fechaFormateada = formatearFecha(fechaCaducidad, "yyyy/MM/dd");
  
  if (!vencimientos[fechaFormateada]) {
    vencimientos[fechaFormateada] = {
      cantidad: 0,
      detalle: [],
      resumen: {}
    };
  }
  
  vencimientos[fechaFormateada].cantidad++;
  
  const sku = row[indicesInventario.sku] || "";
  const descripcion = row[indicesInventario.descripcion] || "";
  const lpn = row[indicesInventario.lpn] || "";
  const cajas = Number(row[indicesInventario.cajas]) || 0;
  
  // Agregar detalle
  vencimientos[fechaFormateada].detalle.push({
    sku,
    descripcion,
    lpn,
    ubicacion: location
  });
  
  // Actualizar resumen
  if (!vencimientos[fechaFormateada].resumen[sku]) {
    vencimientos[fechaFormateada].resumen[sku] = {
      descripcion,
      cantidad: 0,
      cajas: 0
    };
  }
  
  vencimientos[fechaFormateada].resumen[sku].cantidad++;
  vencimientos[fechaFormateada].resumen[sku].cajas += cajas;
}

/**
 * Prepara datos para visualización
 */
function prepararDatosVisualizacion(ubicaciones) {
  const datosUbicaciones = {
    ocupacion: {},
    vencimiento: {}
  };
  
  console.log(`Preparando datos de visualización para ${Object.keys(ubicaciones).length} ubicaciones`);
  
  let contadorLog = 0;
  Object.entries(ubicaciones).forEach(([id, ubicacion]) => {
    const capacidadMaxima = ubicacion.capacidad_maxima || 0;
    const area = ubicacion.area || "";
    const utilizado = ubicacion.utilizado || 0;
    
    const datosComunes = {
      area,
      capacidad_maxima: capacidadMaxima,
      utilizado,
      niveles: ubicacion.niveles || {}
    };
    
    // Datos para ocupación
    let porcentajeOcupacion = 0;
    if (utilizado > 0) {
      porcentajeOcupacion = capacidadMaxima > 0 ? (utilizado / capacidadMaxima) * 100 : 100;
    }
    
    datosUbicaciones.ocupacion[id] = {
      ...datosComunes,
      porcentaje: porcentajeOcupacion.toFixed(1),
      color: porcentajeOcupacion <= 0 ? "#2D572C" : generarColorOcupacion(porcentajeOcupacion)
    };
    
    // Log para debugging (solo las primeras 3 ubicaciones)
    if (contadorLog < 3) {
      console.log(`Ubicación ${id}: utilizado=${utilizado}, capacidad=${capacidadMaxima}, porcentaje=${porcentajeOcupacion.toFixed(1)}%, color=${datosUbicaciones.ocupacion[id].color}`);
      contadorLog++;
    }
    
    // Datos para vencimiento
    const infoVidaUtil = obtenerCategoriasVidaUtil(ubicacion);
    const porcentajeVidaUtil = calcularPorcentajeVidaUtil(ubicacion) || 0;
    
    datosUbicaciones.vencimiento[id] = {
      ...datosComunes,
      porcentaje: porcentajeVidaUtil.toFixed(1),
      color: utilizado <= 0 ? "#666666" : generarColorVidaUtil(porcentajeVidaUtil),
      categoriasVidaUtil: infoVidaUtil.categorias,
      conteoVidaUtil: infoVidaUtil.conteo,
      multipleVidasUtiles: infoVidaUtil.categorias.length > 1
    };
  });
  
  return datosUbicaciones;
}

/**
 * Obtiene categorías de vida útil
 */
function obtenerCategoriasVidaUtil(ubicacion) {
  if (!ubicacion.niveles) return { categorias: [], conteo: {}, total: 0 };
  
  const conteoCategoria = {
    critico: 0,
    bajo: 0,
    medio: 0,
    alto: 0
  };
  
  let totalProductos = 0;
  
  for (const nivelId in ubicacion.niveles) {
    const nivel = ubicacion.niveles[nivelId];
    
    if (nivel && nivel.skus) {
      for (const skuId in nivel.skus) {
        const skuInfo = nivel.skus[skuId];
        if (skuInfo && typeof skuInfo.vida_util_porc === "number") {
          const vidaUtil = skuInfo.vida_util_porc;
          const pallets = Number(skuInfo.pallets) || 1;
          
          totalProductos += pallets;
          
          if (vidaUtil <= 25) conteoCategoria.critico += pallets;
          else if (vidaUtil <= 50) conteoCategoria.bajo += pallets;
          else if (vidaUtil <= 75) conteoCategoria.medio += pallets;
          else conteoCategoria.alto += pallets;
        }
      }
    }
  }
  
  const categoriasOrdenadas = Object.keys(conteoCategoria)
    .filter(cat => conteoCategoria[cat] > 0)
    .sort((a, b) => conteoCategoria[b] - conteoCategoria[a]);
  
  return {
    categorias: categoriasOrdenadas,
    conteo: conteoCategoria,
    total: totalProductos
  };
}

/**
 * Procesa archivo de movimientos (Excel o CSV) para análisis de frecuencia
 */
function procesarMovimientosExcel(blob, periodoInicio, periodoFin, tipoAnalisis = 'ambos') {
  const startTime = Date.now();
  console.log(`Iniciando procesamiento de movimientos: ${periodoInicio} - ${periodoFin}`);
  
  try {
    const config = getConfig();
    const carpeta = DriveApp.getFolderById(config.SUBCARPETA_EXCEL_ID);
    
    let data;
    let archivoTemp = null;
    let spreadsheet = null;
    
    // Detectar tipo de archivo por el nombre
    const fileName = blob.getName() || '';
    const isCSV = fileName.toLowerCase().endsWith('.csv');
    
    if (isCSV) {
      // Procesar archivo CSV
      console.log('Procesando archivo CSV...');
      data = procesarArchivoCSV(blob);
    } else {
      // Procesar archivo Excel
      console.log('Procesando archivo Excel...');
      const nombreTemp = `movimientos_temp_${Date.now()}.xlsx`;
      archivoTemp = carpeta.createFile(blob.setName(nombreTemp));
      
      // Convertir usando la función existente
      spreadsheet = convertirExcelASheets(archivoTemp);
      const sheet = spreadsheet.getSheets()[0];
      data = sheet.getDataRange().getValues();
    }
    
    // Detectar índices de columnas
    const indices = detectarIndicesMovimientos(data[0]);
    console.log('Índices finales para procesamiento:', JSON.stringify(indices));
    
    // Procesar movimientos
    const movimientos = procesarDatosMovimientos(data, indices, periodoInicio, periodoFin, tipoAnalisis);
    
    // Limpiar archivos temporales
    if (archivoTemp) {
      archivoTemp.setTrashed(true);
    }
    if (spreadsheet) {
      DriveApp.getFileById(spreadsheet.getId()).setTrashed(true);
    }
    
    const processingTime = Date.now() - startTime;
    console.log(`Movimientos procesados en ${processingTime}ms`);
    
    return movimientos;
    
  } catch (error) {
    console.error(`Error procesando movimientos: ${error}`);
    throw error;
  }
}

/**
 * Procesa archivo CSV y lo convierte en array de arrays
 */
function procesarArchivoCSV(blob) {
  try {
    const csvContent = blob.getDataAsString('UTF-8');
    const lines = csvContent.split('\n');
    const data = [];
    
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i].trim();
      if (line) {
        // Dividir por comas, pero respetando campos entre comillas
        const row = parseCSVLine(line);
        data.push(row);
      }
    }
    
    console.log(`CSV procesado: ${data.length} filas`);
    return data;
    
  } catch (error) {
    console.error(`Error procesando CSV: ${error}`);
    throw new Error(`Error procesando archivo CSV: ${error.message}`);
  }
}

/**
 * Parsea una línea CSV respetando campos entre comillas
 */
function parseCSVLine(line) {
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
  
  // Limpiar comillas de los campos
  return result.map(field => field.replace(/^"|"$/g, ''));
}

/**
 * Detecta índices de columnas para archivo de movimientos
 */
function detectarIndicesMovimientos(headers) {
  console.log('Headers detectados:', JSON.stringify(headers));
  
  // Normalizar headers para comparación
  const headersNormalizados = headers.map(h => normalizarTexto(h || ''));
  console.log('Headers normalizados:', JSON.stringify(headersNormalizados));
  
  // Detectar si es un archivo con headers válidos
  const tieneHeaders = headers && headers.length > 0 && 
                      (headersNormalizados.some(h => h.includes('almacen') || h.includes('codigo') || h.includes('fecha')));
  
  if (!tieneHeaders || /^[A-Z]+\d+$/.test(headers[0])) {
    console.log('CSV sin encabezados detectado, usando posiciones fijas');
    const indices = {
      warehouse: 0,        // Almacén (PN1500)
      operationCode: 1,    // Código de Operación
      activityCode: 2,     // Código de Actividad  
      transactionDate: 3,  // Fecha
      transactionTime: 4,  // Hora
      quantity: 5,         // Movimientos/Cantidad
      user: 6,            // Usuario
      fullName: 7,        // Nombre Completo
      fromLocation: 8,     // Posición (ubicación real)
      company: 9          // Compañía/Servicio
    };
    console.log('Índices de columnas asignados:', JSON.stringify(indices));
    return indices;
  }
  
  console.log('CSV con encabezados detectado, mapeando columnas...');
  
  // Si hay headers, usar el mapeo mejorado con variantes de codificación
  const columnMapping = {
    warehouse: ['Warehouse', 'Almacen', 'Almacén'],
    operationCode: ['Operation Code', 'Código de Operación', 'Codigo de Operacion', 'codigo de operacion'],
    activityCode: ['Activity Code', 'Código de Actividad', 'Codigo de Actividad', 'codigo de actividad'],
    transactionDate: ['Transaction Date', 'Fecha de Transacción', 'Fecha de Transaccion', 'fecha de transaccion'],
    transactionTime: ['Transaction Time', 'Hora de Transacción', 'Hora de Transaccion', 'hora de transaccion'],
    quantity: ['Quantity', 'Cantidad', 'Movimientos', 'movimientos'],
    user: ['Username', 'Usuario', 'usuario'],
    fullName: ['Full Name', 'Nombre Completo', 'nombre completo'],
    fromLocation: ['From Location', 'Posición', 'Posicion', 'posicion'],
    company: ['Company', 'Compañía', 'Compania', 'compania'],
    dateFrom: ['Date From', 'Desde (DDMMAAAA)', 'Desde', 'desde'],
    dateTo: ['Date To', 'Hasta (DDMMAAAA)', 'Hasta', 'hasta'],
    toLocation: ['To Location', 'Hacia Ubicación'],
    lpn: ['LPN'],
    itemNumber: ['Item Number', 'Número Artículo', 'SKU']
  };
  
  const indices = {};
  
  for (const [key, possibleNames] of Object.entries(columnMapping)) {
    indices[key] = -1; // Inicializar como no encontrado
    
    for (let i = 0; i < headers.length; i++) {
      const headerNormalizado = headersNormalizados[i];
      const headerOriginal = (headers[i] || '').toLowerCase().trim();
      
      for (const nombre of possibleNames) {
        const nombreNormalizado = normalizarTexto(nombre);
        
        // Buscar coincidencia en header normalizado o original
        if (headerNormalizado.includes(nombreNormalizado) || 
            headerOriginal.includes(nombreNormalizado) ||
            headerOriginal.includes(nombre.toLowerCase())) {
          indices[key] = i;
          console.log(`Columna "${key}" encontrada en posición ${i}: "${headers[i]}" -> "${nombre}"`);
          break;
        }
      }
      
      if (indices[key] !== -1) break; // Si encontró la columna, pasar a la siguiente
    }
    
    if (indices[key] === -1) {
      console.warn(`Columna "${key}" no encontrada en headers`);
    }
  }
  
  console.log('Índices finales detectados:', JSON.stringify(indices));
  return indices;
}

/**
 * Procesa datos de movimientos y calcula frecuencias
 */
function procesarDatosMovimientos(data, indices, periodoInicio, periodoFin, tipoAnalisis) {
  const movimientosPorUbicacion = {};
  const fechaInicio = new Date(periodoInicio);
  const fechaFin = new Date(periodoFin);
  
  console.log('Procesando datos con índices:', JSON.stringify(indices));
  console.log('Período:', periodoInicio, '-', periodoFin);
  console.log('Total filas:', data.length);
  
  // Procesar cada fila de movimientos
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    if (i <= 5) { // Log primeras filas para debug
      console.log(`Fila ${i}:`, JSON.stringify(row));
    }
    
    // Verificar fecha en período
    const fechaTransaccion = new Date(row[indices.transactionDate]);
    if (fechaTransaccion < fechaInicio || fechaTransaccion > fechaFin) {
      if (i <= 5) console.log(`Fila ${i} fuera del período:`, fechaTransaccion);
      continue;
    }
    
    // Usar almacén como ubicación si posición está vacía
    let location = row[indices.fromLocation] || row[indices.warehouse];
    if (!location || location.toString().trim() === '') {
      if (i <= 5) console.log(`Fila ${i} sin ubicación válida`);
      continue;
    }
    
    const operationCode = row[indices.operationCode];
    const activityCode = row[indices.activityCode];
    const quantity = Number(row[indices.quantity]) || 1;
    
    if (i <= 5) { // Debug para primeras filas
      console.log(`Fila ${i}: ${JSON.stringify(row)}`);
      console.log(`Índices: operationCode=${indices.operationCode}, activityCode=${indices.activityCode}`);
      console.log(`Valores extraídos: OpCode="${operationCode}", ActCode="${activityCode}"`);
    }
    
    // Determinar si es entrada o salida basado en el código de operación/actividad
    let tipoMovimiento = determinarTipoMovimiento(operationCode, activityCode);
    
    if (i <= 5) { // Debug para primeras filas
      console.log(`Fila ${i} - OpCode: "${operationCode}", ActCode: "${activityCode}", Tipo: "${tipoMovimiento}"`);
    }
    
    // Si no podemos determinar el tipo, continuar
    if (!tipoMovimiento) {
      if (i <= 5) console.log(`Fila ${i} - Tipo de movimiento no determinado, saltando`);
      continue;
    }
    
    // Procesar según el tipo de análisis solicitado
    if ((tipoMovimiento === 'salida' && (tipoAnalisis === 'salidas' || tipoAnalisis === 'ambos')) ||
        (tipoMovimiento === 'entrada' && (tipoAnalisis === 'entradas' || tipoAnalisis === 'ambos'))) {
      
      const idRack = extraerIDSimple(location);
      if (!movimientosPorUbicacion[idRack]) {
        movimientosPorUbicacion[idRack] = {
          salidas: 0,
          entradas: 0,
          detalleMovimientos: []
        };
      }
      
      // Incrementar contador según tipo
      if (tipoMovimiento === 'salida') {
        movimientosPorUbicacion[idRack].salidas += quantity;
      } else {
        movimientosPorUbicacion[idRack].entradas += quantity;
      }
      
      // Agregar detalle si es necesario
      if (movimientosPorUbicacion[idRack].detalleMovimientos.length < 100) {
        movimientosPorUbicacion[idRack].detalleMovimientos.push({
          tipo: tipoMovimiento,
          fecha: formatearFecha(fechaTransaccion, 'dd/MM/yyyy'),
          cantidad: quantity,
          operacion: operationCode,
          actividad: activityCode,
          usuario: row[indices.user] || '',
          almacen: row[indices.warehouse] || ''
        });
      }
    }
  }
  
  // Log resultados finales
  console.log('Movimientos procesados por ubicación:', Object.keys(movimientosPorUbicacion).length);
  console.log('Primeras ubicaciones:', Object.keys(movimientosPorUbicacion).slice(0, 5));
  
  // Calcular estadísticas y asignar colores
  return calcularEstadisticasMovimientos(movimientosPorUbicacion);
}

/**
 * Normaliza texto para manejar problemas de codificación
 */
function normalizarTexto(texto) {
  if (!texto) return '';
  
  let normalizado = texto.toLowerCase().trim();
  
  // Reemplazar caracteres con problemas de codificación
  const reemplazos = {
    'ã³': 'ó',
    'ã¡': 'á', 
    'ã©': 'é',
    'ã­': 'í',
    'ãº': 'ú',
    'ã±': 'ñ',
    'Ã³': 'ó',
    'Ã¡': 'á',
    'Ã©': 'é', 
    'Ã­': 'í',
    'Ãº': 'ú',
    'Ã±': 'ñ'
  };
  
  for (const [mal, bien] of Object.entries(reemplazos)) {
    normalizado = normalizado.replace(new RegExp(mal, 'g'), bien);
  }
  
  return normalizado;
}

/**
 * Determina si un movimiento es entrada o salida basado en códigos
 */
function determinarTipoMovimiento(operationCode, activityCode) {
  // Normalizar códigos para comparación y manejar codificación
  const opCode = normalizarTexto(operationCode);
  const actCode = normalizarTexto(activityCode);
  
  console.log(`Clasificando movimiento: Op="${opCode}", Act="${actCode}"`);
  
  // Si ambos códigos están vacíos, no se puede clasificar
  if (!opCode && !actCode) {
    console.log('Ambos códigos están vacíos, no se puede clasificar');
    return null;
  }
  
  // Definir combinaciones específicas que indican ENTRADA
  const combinacionesEntrada = [
    // Transferencia + Movimientos de inventario
    { op: 'transferencia de carga no dirigida', act: 'movimiento de inventario completo no dirigido' },
    { op: 'transferencia de carga no dirigida', act: 'movimiento de inventario completo no dirigido con reserva' },
    { op: 'transferencia de carga no dirigida', act: 'movimiento de inventario parcial no dirigido' },
    
    // Identificación + Recepción/Movimientos
    { op: 'identificación no dirigida', act: 'recepción' },
    { op: 'identificación no dirigida', act: 'movimiento de inventario completo no dirigido' },
    { op: 'identificación no dirigida', act: 'movimiento de inventario completo no dirigido con reserva' },
    
    // Reabastecimiento
    { op: 'reabast de palés', act: 'reabastecimiento de palés' }
  ];
  
  // Definir combinaciones específicas que indican SALIDA
  const combinacionesSalida = [
    // Surtido + Carga
    { op: 'surtido pallet doble', act: 'carga de tráiler' },
    { op: 'surtido pallet doble', act: 'carga fluida' },
    
    // Transferencia + Carga (estas son salidas, no entradas)
    { op: 'transferencia de carga no dirigida', act: 'carga de tráiler' },
    { op: 'transferencia de carga no dirigida', act: 'carga fluida' },
    
    // Asignación + Carga
    { op: 'asignación de trabajo', act: 'carga de tráiler' },
    { op: 'asignación de trabajo', act: 'carga fluida' }
  ];
  
  // Verificar combinaciones de entrada
  for (const combo of combinacionesEntrada) {
    if (opCode === combo.op && actCode === combo.act) {
      console.log(`Clasificado como ENTRADA por combinación: "${combo.op}" + "${combo.act}"`);
      return 'entrada';
    }
  }
  
  // Verificar combinaciones de salida
  for (const combo of combinacionesSalida) {
    if (opCode === combo.op && actCode === combo.act) {
      console.log(`Clasificado como SALIDA por combinación: "${combo.op}" + "${combo.act}"`);
      return 'salida';
    }
  }
  
  // Reglas por palabras clave para casos no cubiertos por combinaciones exactas
  
  // Palabras clave que indican SALIDA
  const palabrasClavesSalida = ['carga de tráiler', 'carga fluida', 'surtido', 'pick', 'shipment'];
  for (const palabra of palabrasClavesSalida) {
    if (opCode.includes(palabra) || actCode.includes(palabra)) {
      console.log(`Clasificado como SALIDA por palabra clave: "${palabra}"`);
      return 'salida';
    }
  }
  
  // Palabras clave que indican ENTRADA
  const palabrasClaveEntrada = ['movimiento de inventario', 'recepción', 'reabastecimiento', 'identificación'];
  for (const palabra of palabrasClaveEntrada) {
    if (opCode.includes(palabra) || actCode.includes(palabra)) {
      console.log(`Clasificado como ENTRADA por palabra clave: "${palabra}"`);
      return 'entrada';
    }
  }
  
  console.log(`No se pudo clasificar el movimiento: Op="${opCode}", Act="${actCode}"`);
  return null;
}

/**
 * Calcula estadísticas y asigna colores según frecuencia de movimientos
 */
function calcularEstadisticasMovimientos(movimientos) {
  // Obtener todos los valores de salidas y entradas
  const todasSalidas = Object.values(movimientos).map(m => m.salidas).filter(s => s > 0);
  const todasEntradas = Object.values(movimientos).map(m => m.entradas).filter(e => e > 0);
  
  // Calcular percentiles para salidas
  const percentilesSalidas = calcularPercentiles(todasSalidas);
  const percentilesEntradas = calcularPercentiles(todasEntradas);
  
  // Asignar colores a cada ubicación
  for (const [ubicacion, datos] of Object.entries(movimientos)) {
    // Color para salidas
    if (datos.salidas > 0) {
      datos.colorSalidas = generarColorMovimiento(datos.salidas, percentilesSalidas);
      datos.percentilSalidas = obtenerPercentil(datos.salidas, percentilesSalidas);
    } else {
      datos.colorSalidas = '#2D572C'; // Verde oscuro para sin movimiento
      datos.percentilSalidas = 0;
    }
    
    // Color para entradas
    if (datos.entradas > 0) {
      datos.colorEntradas = generarColorMovimiento(datos.entradas, percentilesEntradas);
      datos.percentilEntradas = obtenerPercentil(datos.entradas, percentilesEntradas);
    } else {
      datos.colorEntradas = '#2D572C'; // Verde oscuro para sin movimiento
      datos.percentilEntradas = 0;
    }
  }
  
  return {
    movimientos: movimientos,
    estadisticas: {
      totalUbicaciones: Object.keys(movimientos).length,
      percentilesSalidas: percentilesSalidas,
      percentilesEntradas: percentilesEntradas,
      promedioSalidas: todasSalidas.length > 0 ? todasSalidas.reduce((a, b) => a + b, 0) / todasSalidas.length : 0,
      promedioEntradas: todasEntradas.length > 0 ? todasEntradas.reduce((a, b) => a + b, 0) / todasEntradas.length : 0
    }
  };
}

/**
 * Calcula percentiles de un array de valores
 */
function calcularPercentiles(valores) {
  if (valores.length === 0) return { p25: 0, p50: 0, p75: 0, p100: 0 };
  
  const sorted = valores.sort((a, b) => a - b);
  const len = sorted.length;
  
  return {
    p25: sorted[Math.floor(len * 0.25)],
    p50: sorted[Math.floor(len * 0.50)],
    p75: sorted[Math.floor(len * 0.75)],
    p100: sorted[len - 1]
  };
}

/**
 * Obtiene el percentil de un valor
 */
function obtenerPercentil(valor, percentiles) {
  if (valor <= percentiles.p25) return 25;
  if (valor <= percentiles.p50) return 50;
  if (valor <= percentiles.p75) return 75;
  return 100;
}

/**
 * Genera color según frecuencia de movimiento
 */
function generarColorMovimiento(frecuencia, percentiles) {
  if (frecuencia <= percentiles.p25) {
    return '#00FF00'; // Verde - Bajo movimiento
  } else if (frecuencia <= percentiles.p50) {
    return '#FFFF00'; // Amarillo - Movimiento medio
  } else if (frecuencia <= percentiles.p75) {
    return '#FF8800'; // Naranja - Movimiento alto
  } else {
    return '#FF0000'; // Rojo - Movimiento muy alto
  }
}

/**
 * Calcula porcentaje de vida útil ponderado
 */
function calcularPorcentajeVidaUtil(ubicacion) {
  if (!ubicacion.niveles) return 0;
  
  let totalProductos = 0;
  let sumaPorcentajes = 0;
  
  for (const nivelId in ubicacion.niveles) {
    const nivel = ubicacion.niveles[nivelId];
    
    if (nivel && nivel.skus) {
      for (const skuId in nivel.skus) {
        const skuInfo = nivel.skus[skuId];
        if (skuInfo && typeof skuInfo.vida_util_porc === "number") {
          const vidaUtil = Number(skuInfo.vida_util_porc) || 0;
          const pallets = Number(skuInfo.pallets) || 1;
          
          sumaPorcentajes += vidaUtil * pallets;
          totalProductos += pallets;
        }
      }
    }
  }
  
  return totalProductos > 0 ? sumaPorcentajes / totalProductos : 0;
}

/**
 * Genera mapa de SKU
 */
function generarMapaSKU(ubicaciones) {
  const mapaSKU = {};
  
  Object.entries(ubicaciones).forEach(([id, ubicacion]) => {
    if (ubicacion.niveles) {
      Object.entries(ubicacion.niveles).forEach(([nivel, infoNivel]) => {
        if (infoNivel.skus && typeof infoNivel.skus === "object") {
          Object.entries(infoNivel.skus).forEach(([codigoSku, skuInfo]) => {
            if (!mapaSKU[codigoSku]) {
              mapaSKU[codigoSku] = {
                descripcion: skuInfo.descripcion || "",
                ubicaciones: []
              };
            }
            
            mapaSKU[codigoSku].ubicaciones.push({
              ubicacion: extraerIDSimple(nivel),
              cantidad: skuInfo.pallets || 0
            });
          });
        }
      });
    }
  });
  
  return mapaSKU;
}

/**
 * Procesa datos de movimientos desde archivo de inventario CSV
 */
function procesarDatosMovimientosInventario(data, indices, periodoInicio, periodoFin, tipoAnalisis) {
  const movimientosPorUbicacion = {};
  const fechaInicio = new Date(periodoInicio);
  const fechaFin = new Date(periodoFin);
  
  console.log('Procesando movimientos de inventario con índices:', JSON.stringify(indices));
  console.log('Período:', periodoInicio, '-', periodoFin);
  console.log('Total filas de inventario:', data.length);
  
  // Procesar cada fila de movimientos del inventario
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    if (i <= 5) { // Log primeras filas para debug
      console.log(`Fila inventario ${i}:`, JSON.stringify(row));
    }
    
    // Obtener fecha de transacción
    let fechaTransaccion = null;
    if (indices.transactionDate >= 0 && row[indices.transactionDate]) {
      try {
        const fechaStr = row[indices.transactionDate].toString().trim();
        // Intentar parsear diferentes formatos de fecha
        if (fechaStr.includes('/')) {
          // Formato DD/MM/YYYY HH:mm:ss
          const partes = fechaStr.split(' ');
          const fecha = partes[0]; // DD/MM/YYYY
          const [dia, mes, año] = fecha.split('/');
          fechaTransaccion = new Date(año, mes - 1, dia);
        } else {
          fechaTransaccion = new Date(fechaStr);
        }
      } catch (error) {
        console.warn(`Error parseando fecha en fila ${i}: ${row[indices.transactionDate]}`);
        continue;
      }
    }
    
    // Verificar si la fecha está en el período especificado
    if (!fechaTransaccion || fechaTransaccion < fechaInicio || fechaTransaccion > fechaFin) {
      if (i <= 5) console.log(`Fila ${i} fuera del período:`, fechaTransaccion);
      continue;
    }
    
    // Obtener ubicaciones de origen y destino
    const ubicacionOrigen = indices.fromLocation >= 0 ? row[indices.fromLocation] : '';
    const ubicacionDestino = indices.toLocation >= 0 ? row[indices.toLocation] : '';
    const areaOrigen = indices.fromArea >= 0 ? row[indices.fromArea] : '';
    const areaDestino = indices.toArea >= 0 ? row[indices.toArea] : '';
    
    // Determinar la ubicación principal a usar
    let ubicacionPrincipal = ubicacionOrigen || ubicacionDestino;
    if (!ubicacionPrincipal || ubicacionPrincipal.toString().trim() === '') {
      if (i <= 5) console.log(`Fila ${i} sin ubicación válida`);
      continue;
    }
    
    // Obtener información del movimiento
    const operacion = indices.operation >= 0 ? row[indices.operation] : '';
    const actividad = indices.activity >= 0 ? row[indices.activity] : '';
    const cantidad = Number(row[indices.quantity]) || 1;
    const articulo = indices.item >= 0 ? row[indices.item] : '';
    
    if (i <= 5) { // Debug para primeras filas
      console.log(`Fila ${i}: Operación="${operacion}", Actividad="${actividad}", Cantidad=${cantidad}`);
      console.log(`Ubicaciones: Origen="${ubicacionOrigen}", Destino="${ubicacionDestino}"`);
    }
    
    // Determinar tipo de movimiento basado en ubicaciones y operación
    let tipoMovimiento = determinarTipoMovimientoInventario(
      operacion, 
      actividad, 
      ubicacionOrigen, 
      ubicacionDestino,
      areaOrigen,
      areaDestino
    );
    
    if (i <= 5) { // Debug para primeras filas
      console.log(`Fila ${i} - Tipo movimiento determinado: "${tipoMovimiento}"`);
    }
    
    // Si no podemos determinar el tipo, continuar
    if (!tipoMovimiento) {
      if (i <= 5) console.log(`Fila ${i} - Tipo de movimiento no determinado, saltando`);
      continue;
    }
    
    // Procesar según el tipo de análisis solicitado
    if ((tipoMovimiento === 'salida' && (tipoAnalisis === 'salidas' || tipoAnalisis === 'ambos')) ||
        (tipoMovimiento === 'entrada' && (tipoAnalisis === 'entradas' || tipoAnalisis === 'ambos'))) {
      
      const idRack = extraerIDSimple(ubicacionPrincipal);
      if (!movimientosPorUbicacion[idRack]) {
        movimientosPorUbicacion[idRack] = {
          salidas: 0,
          entradas: 0,
          detalleMovimientos: []
        };
      }
      
      // Incrementar contador según tipo
      if (tipoMovimiento === 'salida') {
        movimientosPorUbicacion[idRack].salidas += cantidad;
      } else {
        movimientosPorUbicacion[idRack].entradas += cantidad;
      }
      
      // Agregar detalle si es necesario (limitar a 100 por ubicación)
      if (movimientosPorUbicacion[idRack].detalleMovimientos.length < 100) {
        movimientosPorUbicacion[idRack].detalleMovimientos.push({
          tipo: tipoMovimiento,
          fecha: formatearFecha(fechaTransaccion, 'dd/MM/yyyy'),
          cantidad: cantidad,
          operacion: operacion,
          actividad: actividad,
          articulo: articulo,
          ubicacionOrigen: ubicacionOrigen,
          ubicacionDestino: ubicacionDestino,
          usuario: indices.user >= 0 ? row[indices.user] : ''
        });
      }
    }
  }
  
  // Log resultados finales
  console.log('Movimientos de inventario procesados por ubicación:', Object.keys(movimientosPorUbicacion).length);
  console.log('Primeras ubicaciones:', Object.keys(movimientosPorUbicacion).slice(0, 5));
  
  // Calcular estadísticas y asignar colores
  return calcularEstadisticasMovimientos(movimientosPorUbicacion);
}

/**
 * Determina el tipo de movimiento basado en datos de inventario
 */
function determinarTipoMovimientoInventario(operacion, actividad, ubicacionOrigen, ubicacionDestino, areaOrigen, areaDestino) {
  // Normalizar textos para comparación
  const op = (operacion || '').toString().toLowerCase().trim();
  const act = (actividad || '').toString().toLowerCase().trim();
  const origen = (ubicacionOrigen || '').toString().trim();
  const destino = (ubicacionDestino || '').toString().trim();
  
  // Reglas específicas para movimientos de inventario
  
  // Si hay ubicación de destino pero no origen, es entrada
  if (destino && !origen) {
    return 'entrada';
  }
  
  // Si hay ubicación de origen pero no destino, es salida
  if (origen && !destino) {
    return 'salida';
  }
  
  // Basado en operación y actividad
  if (op.includes('recepcion') || op.includes('recepción') || 
      op.includes('ingreso') || op.includes('entrada') ||
      act.includes('recepcion') || act.includes('recepción') ||
      act.includes('ingreso') || act.includes('entrada')) {
    return 'entrada';
  }
  
  if (op.includes('despacho') || op.includes('salida') || op.includes('envio') || op.includes('envío') ||
      act.includes('despacho') || act.includes('salida') || act.includes('envio') || act.includes('envío')) {
    return 'salida';
  }
  
  // Movimientos internos o cambios de atributo se consideran como movimientos internos
  if (op.includes('cambio') || op.includes('movimiento') || op.includes('traslado') ||
      act.includes('cambio') || act.includes('movimiento') || act.includes('traslado')) {
    // Para movimientos internos, considerar como movimiento general (entrada)
    return 'entrada';
  }
  
  // Si no se puede determinar, asumir como movimiento interno (entrada)
  return 'entrada';
}