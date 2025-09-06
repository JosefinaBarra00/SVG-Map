/**
 * Procesa un archivo Excel enviado como base64
 * @param {string} base64data - Datos del archivo en formato base64
 * @param {string} fileName - Nombre del archivo
 * @return {Object} - Resultado del procesamiento
 */

function procesarArchivoExcelADatos(
  blob,
  customTimestamp,
  useExistingFile = false
) {
  try {
    console.log("Procesando archivo excel", customTimestamp);

    if (useExistingFile) {
      console.log("Usando archivo existente sin crear uno nuevo");

      try {
        // Intentar convertir a Google Sheets para extraer datos correctamente
        const driveFile = DriveApp.getFileById(
          blob.getId ? blob.getId() : null
        );

        if (!driveFile) {
          throw new Error("No se pudo obtener referencia al archivo existente");
        }

        // Usar Drive API para abrir sin convertir
        let inventarioData;

        try {
          // Intentar abrir como Google Sheets si ya es un archivo de Google Sheets
          const ss = SpreadsheetApp.openById(driveFile.getId());
          const sheet = ss.getSheets()[0];
          inventarioData = sheet.getDataRange().getValues();
          console.log(
            "Datos extraídos directamente de Google Sheets existente"
          );
        } catch (e) {
          // Si no es un Google Sheet, intentar convertir temporalmente
          const resource = {
            title: driveFile.getName() + " (Temp)",
            mimeType: MimeType.GOOGLE_SHEETS,
            parents: [{ id: SUBCARPETA_EXCEL_ID }],
          };

          const tempFile = Drive.Files.insert(resource, driveFile.getBlob());
          console.log(
            "Archivo convertido temporalmente para extracción de datos"
          );

          // Dar tiempo para que termine la conversión
          Utilities.sleep(2000);

          // Abrir la hoja y extraer datos
          const ss = SpreadsheetApp.openById(tempFile.getId());
          const sheet = ss.getSheets()[0];
          inventarioData = sheet.getDataRange().getValues();

          // Borrar el archivo temporal
          DriveApp.getFileById(tempFile.getId()).setTrashed(true);
          console.log("Archivo temporal eliminado después de extracción");
        }

        if (inventarioData && inventarioData.length > 1) {
          return { inventario: inventarioData };
        } else {
          throw new Error("No se pudieron extraer datos suficientes");
        }
      } catch (error) {
        console.error("Error al extraer datos del archivo existente:", error);
        throw error;
      }
    }

    // Usar la fecha personalizada si existe, o la fecha actual
    let fecha;
    if (customTimestamp) {
      console.log("Usando timestamp personalizado:", customTimestamp);
      try {
        const timestampNum =
          typeof customTimestamp === "number"
            ? customTimestamp
            : parseInt(customTimestamp);

        if (!isNaN(timestampNum)) {
          fecha = new Date(timestampNum);
          console.log("Fecha decodificada:", fecha);
        } else {
          fecha = new Date();
          console.log("Timestamp inválido, usando fecha actual:", fecha);
        }
      } catch (e) {
        fecha = new Date();
        console.log("Error al parsear timestamp, usando fecha actual:", fecha);
      }
    } else {
      fecha = new Date();
      console.log("Sin timestamp, usando fecha actual:", fecha);
    }

    // Generar nombre definitivo con formato estándar, asegurando que se usa la hora correcta
    const dia = fecha.getDate().toString().padStart(2, "0");
    const mes = (fecha.getMonth() + 1).toString().padStart(2, "0");
    const anio = fecha.getFullYear();
    const hora = fecha.getHours().toString().padStart(2, "0");
    const minutos = fecha.getMinutes().toString().padStart(2, "0");

    console.log("Componentes de fecha para nombre archivo:", {
      dia,
      mes,
      anio,
      hora,
      minutos,
    });

    const inventarioFileName = `Inventario_${dia}_${mes}_${anio}_${hora}_${minutos}.xlsx`;
    console.log("Nombre de archivo generado:", inventarioFileName);

    // Obtener la carpeta de destino
    const carpetaTemp = DriveApp.getFolderById(SUBCARPETA_EXCEL_ID);

    // Eliminar SOLO archivos temporales y convertidos, preservar los Inventario_*
    // excepto el más reciente si estamos creando uno nuevo
    const allFiles = carpetaTemp.getFiles();
    let archivoInventarioAnterior = null;
    let fechaArchivoAnterior = new Date(0);

    // Primera pasada: identificar el archivo Inventario más reciente y borrar temporales
    while (allFiles.hasNext()) {
      const file = allFiles.next();
      const fileName = file.getName();

      if (
        fileName.startsWith("temp_") ||
        fileName.includes("(Convertido)") ||
        fileName.includes("(Procesado)")
      ) {
        // Eliminar archivos temporales y convertidos
        file.setTrashed(true);
      } else if (fileName.startsWith("Inventario_")) {
        // Guardar referencia al más reciente para decidir después
        const fileDate = file.getLastUpdated();
        if (fileDate > fechaArchivoAnterior) {
          fechaArchivoAnterior = fileDate;
          archivoInventarioAnterior = file;
        }
      }
    }

    // Si estamos creando un archivo nuevo de Inventario, borrar el anterior
    /* if (archivoInventarioAnterior !== null) {
      archivoInventarioAnterior.setTrashed(true);
      console.log("Archivo anterior eliminado: " + archivoInventarioAnterior.getName());
    } */

    // Guardar el archivo en Drive para procesamiento
    const archivoTemp = carpetaTemp.createFile(
      blob.setName(inventarioFileName)
    );
    console.log("Archivo guardado en Drive: " + archivoTemp.getName());

    // Esperar a que Drive procese el archivo
    Utilities.sleep(3000);

    try {
      // Intentar convertir a Google Sheets para extraer datos correctamente
      const driveFile = DriveApp.getFileById(archivoTemp.getId());
      const resource = {
        title: inventarioFileName + " (Procesado)",
        mimeType: MimeType.GOOGLE_SHEETS,
        parents: [{ id: SUBCARPETA_EXCEL_ID }],
      };

      // Usar Drive API para convertir
      const newFile = Drive.Files.insert(resource, driveFile.getBlob());
      console.log("Archivo convertido a Google Sheets: " + newFile.getId());

      // Dar tiempo para que termine la conversión
      Utilities.sleep(3000);

      // Abrir la hoja y extraer datos
      const ss = SpreadsheetApp.openById(newFile.getId());
      const sheet = ss.getSheets()[0];
      const range = sheet.getDataRange();
      const inventarioData = range.getValues();

      console.log(
        "Datos extraídos exitosamente. Filas: " + inventarioData.length
      );

      // Borrar el archivo convertido para evitar duplicados
      DriveApp.getFileById(newFile.getId()).setTrashed(true);

      // Verificar datos
      if (inventarioData && inventarioData.length > 1) {
        console.log(
          "Primera fila de datos:",
          JSON.stringify(inventarioData[0])
        );
        console.log(
          "Segunda fila de datos:",
          JSON.stringify(inventarioData[1])
        );

        return {
          inventario: inventarioData,
        };
      } else {
        console.log("La hoja no contiene suficientes datos (solo encabezados)");
        throw new Error("La hoja convertida no contiene suficientes datos");
      }
    } catch (conversionError) {
      console.log(
        "Error en la conversión o extracción de datos: " + conversionError
      );

      // Intentar con otro método de extracción
      try {
        // Método alternativo: usar XlsxApp si está disponible
        if (typeof XlsxApp !== "undefined") {
          const xlsFile = XlsxApp.open(archivoTemp.getBlob());
          const sheet = xlsFile.getSheets()[0];
          const inventarioData = sheet.getDataRange().getValues();

          console.log(
            "Datos extraídos con XlsxApp. Filas: " + inventarioData.length
          );

          if (inventarioData && inventarioData.length > 1) {
            return { inventario: inventarioData };
          }
        }

        // Si no podemos extraer directamente, usar los datos del inventario existente
        console.log("Usando método de respaldo para obtener datos");
        const sheet =
          SpreadsheetApp.openById(INVENTARIO_SHEET_ID).getSheetByName(
            SHEET_INVENTARIO
          );
        const inventarioData = sheet.getDataRange().getValues();

        if (inventarioData.length <= 1) {
          throw new Error(
            "Los datos de respaldo están vacíos o solo contienen encabezados"
          );
        }

        console.log(
          "Datos de respaldo obtenidos. Filas: " + inventarioData.length
        );
        return { inventario: inventarioData };
      } catch (backupError) {
        console.error("Error en método de respaldo: " + backupError);
        throw backupError;
      }
    }
  } catch (error) {
    console.error("Error completo en procesamiento: " + error.toString());

    // Intentar un último método de respaldo
    try {
      const sheet =
        SpreadsheetApp.openById(INVENTARIO_SHEET_ID).getSheetByName(
          SHEET_INVENTARIO
        );
      const inventarioData = sheet.getDataRange().getValues();

      if (inventarioData.length > 1) {
        console.log(
          "Usando datos existentes como último recurso. Filas: " +
            inventarioData.length
        );
        return { inventario: inventarioData };
      }
    } catch (finalError) {
      console.error("Error final: " + finalError);
    }

    // Retornar una matriz vacía como última opción
    return {
      inventario: [
        ["Ubicación", "SKU", "Descripción", "LPN", "Cantidad", "Fecha"],
      ],
    };
  }
}

// Función original de extraerIDSimple proporcionada
function extraerIDSimple(idCompleto) {
  if (!idCompleto) return null;

  // Convertir a string y limpiar
  idCompleto = idCompleto.toString().trim();

  // Intenta encontrar patrones como L03-12 o similares (primera parte hasta el segundo guion)
  var match = idCompleto.match(/^([A-Z][0-9]+(?:-[0-9]+)?)/);
  if (match) {
    return match[1];
  }

  // Si el ID contiene guiones, tomar los primeros dos segmentos
  if (idCompleto.includes("-")) {
    var parts = idCompleto.split("-");
    if (parts.length >= 2) {
      return parts[0] + "-" + parts[1];
    }
  }

  // Si no se pudo extraer según los patrones anteriores, devolver el ID original
  // Logger.log("No se pudo extraer ID para: " + idCompleto + ", retornando original");
  return idCompleto;
}

function procesarUbicaciones(dataUbicaciones, indices) {
  const ubicaciones = {};

  // Iniciar desde la fila 1 para evitar encabezados
  for (let i = 1; i < dataUbicaciones.length; i++) {
    const location = dataUbicaciones[i][indices.ubicacion];
    if (!location) continue;

    const idRack = extraerIDSimple(location);
    const area = dataUbicaciones[i][indices.area] || "";

    // Inicializar el rack si no existe
    if (!ubicaciones[idRack]) {
      ubicaciones[idRack] = {
        area: area,
        zona_trabajo: "",
        capacidad_maxima: dataUbicaciones[i][indices.capacidad_maxima] || 0, // Se calculará automáticamente contando los niveles
        utilizado: 0,
        niveles: {},
      };
    }

    // Inicializar nivel si no existe
    if (!ubicaciones[idRack].niveles[location]) {
      ubicaciones[idRack].niveles[location] = {
        utilizado: 0,
        skus: {},
      };
    }

    // Actualizar capacidad_maxima contando los niveles
    if (ubicaciones[idRack].capacidad_maxima === 0) {
      ubicaciones[idRack].capacidad_maxima = Object.keys(
        ubicaciones[idRack].niveles
      ).length;
    }
  }

  return ubicaciones;
}

function procesarUbicacionesMultinivel(
  ubicaciones,
  data,
  dataSKU,
  indicesInventario,
  indicesSKU
) {
  // Inicializar una sola vez
  const vencimientos = {};

  const skuVidaUtilMap = {};
  if (dataSKU && dataSKU.length > 0) {
    for (let i = 1; i < dataSKU.length; i++) {
      if (dataSKU[i] && dataSKU[i][indicesSKU.sku]) {
        skuVidaUtilMap[dataSKU[i][indicesSKU.sku]] =
          Number(dataSKU[i][indicesSKU.vida_util]) || 0;
      }
    }
  }

  // Fecha actual (calculada una sola vez)
  const hoy = new Date();

  // Comenzar desde la segunda fila (índice 1)
  const batchSize = 100;
  for (let i = 1; i < data.length; i += batchSize) {
    const endIdx = Math.min(i + batchSize, data.length);
    procesarLoteUbicaciones(
      data.slice(i, endIdx),
      ubicaciones,
      vencimientos,
      skuVidaUtilMap,
      hoy,
      indicesInventario
    );
  }
  return {
    ubicaciones: ubicaciones,
    vencimientos: vencimientos,
  };
}

function procesarLoteUbicaciones(
  loteData,
  ubicaciones,
  vencimientos,
  skuVidaUtilMap,
  hoy,
  indicesInventario
) {
  for (let j = 0; j < loteData.length; j++) {
    const row = loteData[j];
    const location = row[indicesInventario.ubicacion];
    if (!location) continue; // Saltar filas sin ubicación

    const idRack = extraerIDSimple(location);
    const cantidad = Number(row[indicesInventario.cajas]) > 0 ? 1 : 0;

    // Inicializar o actualizar ubicación
    if (!ubicaciones[idRack]) {
      ubicaciones[idRack] = {
        area: row[indicesInventario.area] || "",
        zona_trabajo: row[indicesInventario.zona_trabajo] || "",
        capacidad_maxima: 1,
        utilizado: cantidad,
        niveles: {},
      };
    } else {
      ubicaciones[idRack].utilizado += cantidad;
    }

    // Inicializar el nivel si no existe
    if (!ubicaciones[idRack].niveles[location]) {
      ubicaciones[idRack].niveles[location] = {
        utilizado: 0,
        skus: {},
      };
    } else if (!ubicaciones[idRack].niveles[location].skus) {
      // Añadir el array lpns si no existe (para compatibilidad con datos existentes)
      ubicaciones[idRack].niveles[location].skus = {};
    }

    const sku = row[indicesInventario.sku]
      ? String(row[indicesInventario.sku])
      : "";

    // Datos de nivel
    if (sku && cantidad > 0) {
      if (!ubicaciones[idRack].niveles[location].skus[sku]) {
        ubicaciones[idRack].niveles[location].skus[sku] = {
          sku: sku,
          descripcion: row[indicesInventario.descripcion]
            ? String(row[indicesInventario.descripcion])
            : "",
          cajas: 0,
          pallets: 0,
          lpns: [], // Array vacío para LPNs
          fechas: [],
        };
      }

      const skuObj = ubicaciones[idRack].niveles[location].skus[sku];

      // Actualizar el total de cajas para este SKU
      const cajas =
        row[indicesInventario.cajas] &&
        !isNaN(Number(row[indicesInventario.cajas]))
          ? Number(row[indicesInventario.cajas])
          : 0;
      skuObj.cajas += cajas;

      // Incrementar contador de pallets para este SKU
      skuObj.pallets += cantidad;

      // Añadir LPN a la lista si no está ya (opcional)
      const lpn = row[indicesInventario.lpn]
        ? String(row[indicesInventario.lpn])
        : "";
      if (lpn && !skuObj.lpns.includes(lpn)) {
        skuObj.lpns.push(lpn);
      }

      let diasRestantes;
      let vidaUtilTotal;

      const fecha_caducidad = new Date(row[indicesInventario.caducidad]);
      const fecha_formateada = Utilities.formatDate(
        fecha_caducidad,
        Session.getScriptTimeZone(),
        "dd/MM/yyyy"
      );

      if (fecha_formateada && !skuObj.fechas.includes(fecha_formateada)) {
        skuObj.fechas.push(fecha_formateada);
      }

      if (row[indicesInventario.diasHastaCaducidad]) {
        diasRestantes = row[indicesInventario.diasHastaCaducidad];
        vidaUtilTotal = row[indicesInventario.perfilAntiguedad];
      } else if (row[indicesInventario.caducidad]) {
        if (fecha_caducidad && !isNaN(fecha_caducidad.getTime())) {
          // Solo calculamos los días restantes si es necesario
          diasRestantes = Math.max(
            0,
            Math.floor((fecha_caducidad - hoy) / (1000 * 60 * 60 * 24))
          );
          vidaUtilTotal = skuVidaUtilMap[sku] || 0;
        }
      }

      if (vidaUtilTotal > 0) {
        // Calculamos y guardamos directamente
        skuObj.vida_util_dias = diasRestantes;
        skuObj.vida_util_porc = Math.min(
          100,
          Math.round((diasRestantes / vidaUtilTotal) * 100)
        );
      }

      // Actualizar el total utilizado para el nivel
      ubicaciones[idRack].niveles[location].utilizado += cantidad;

      // Actualizar capacidad máxima si es necesario
      if (Object.keys(ubicaciones[idRack].niveles).length > 1) {
        ubicaciones[idRack].capacidad_maxima = Object.keys(
          ubicaciones[idRack].niveles
        ).length;
      }
    }

    // Procesar vencimientos si hay datos
    procesarVencimiento(row, location, vencimientos, indicesInventario);
  }
}

// Función auxiliar para procesar un vencimiento
function procesarVencimiento(row, location, vencimientos, indicesInventario) {
  const fecha_caducidad = new Date(row[indicesInventario.caducidad]);
  if (!fecha_caducidad) return;

  // Calcular fecha de vencimiento
  const fechaFormateada = Utilities.formatDate(
    fecha_caducidad,
    Session.getScriptTimeZone(),
    "yyyy/MM/dd"
  );

  // Inicializar si es nuevo
  if (!vencimientos[fechaFormateada]) {
    vencimientos[fechaFormateada] = {
      cantidad: 0,
      detalle: [],
      resumen: {},
    };
  }

  // Incrementar contador
  vencimientos[fechaFormateada].cantidad++;

  // Obtener datos del producto
  const sku = row[indicesInventario.sku] || "";
  const descripcion = row[indicesInventario.descripcion] || "";
  const lpn = row[indicesInventario.lpn] || "";
  const cajas = Number(row[indicesInventario.cajas]) || 0;

  // Añadir al detalle
  vencimientos[fechaFormateada].detalle.push({
    sku: sku,
    descripcion: descripcion,
    lpn: lpn,
    ubicacion: location,
  });

  // Actualizar resumen por SKU
  if (!vencimientos[fechaFormateada].resumen[sku]) {
    vencimientos[fechaFormateada].resumen[sku] = {
      descripcion: descripcion,
      cantidad: 0,
      cajas: 0,
    };
  }
  vencimientos[fechaFormateada].resumen[sku].cantidad++;
  vencimientos[fechaFormateada].resumen[sku].cajas += cajas;
}

function obtenerCategoriasVidaUtil(ubicacion) {
  if (!ubicacion.niveles) return { categorias: [], conteo: {}, total: 0 };

  // Contar la cantidad de productos en cada categoría
  const conteoCategoria = {
    critico: 0,
    bajo: 0,
    medio: 0,
    alto: 0,
  };

  let totalProductos = 0;

  // Recorrer todos los niveles
  for (const nivelId in ubicacion.niveles) {
    const nivel = ubicacion.niveles[nivelId];

    if (nivel && nivel.skus) {
      for (const skuId in nivel.skus) {
        const skuInfo = nivel.skus[skuId];
        if (skuInfo && typeof skuInfo.vida_util_porc === "number") {
          const vidaUtil = skuInfo.vida_util_porc;
          // Asegurarse de contar correctamente los pallets, usando el valor real o defaulteando a 1
          const pallets =
            typeof skuInfo.pallets === "number" && !isNaN(skuInfo.pallets)
              ? Number(skuInfo.pallets)
              : 1;
          totalProductos += pallets;

          if (vidaUtil <= 25) conteoCategoria["critico"] += pallets;
          else if (vidaUtil <= 50) conteoCategoria["bajo"] += pallets;
          else if (vidaUtil <= 75) conteoCategoria["medio"] += pallets;
          else conteoCategoria["alto"] += pallets;
        }
      }
    }
  }

  // Convertir a array y ordenar por cantidad descendente
  const categoriasOrdenadas = Object.keys(conteoCategoria)
    .filter((cat) => conteoCategoria[cat] > 0)
    .sort((a, b) => conteoCategoria[b] - conteoCategoria[a]);

  return {
    categorias: categoriasOrdenadas,
    conteo: conteoCategoria,
    total: totalProductos,
  };
}

function calcularPorcentajeVidaUtil(ubicacion) {
  if (!ubicacion.niveles) return 0;

  let totalProductos = 0;
  let sumaPorcentajes = 0;

  // Recorrer todos los niveles
  for (const nivelId in ubicacion.niveles) {
    const nivel = ubicacion.niveles[nivelId];

    // Nueva estructura con SKUs
    if (nivel && nivel.skus) {
      for (const skuId in nivel.skus) {
        const skuInfo = nivel.skus[skuId];
        if (skuInfo && typeof skuInfo.vida_util_porc === "number") {
          const vidaUtil = isNaN(skuInfo.vida_util_porc)
            ? 0
            : Number(skuInfo.vida_util_porc);
          // Asegurarse de contar correctamente los pallets para la ponderación
          const pallets =
            typeof skuInfo.pallets === "number" && !isNaN(skuInfo.pallets)
              ? Number(skuInfo.pallets)
              : 1;
          sumaPorcentajes += vidaUtil * pallets;
          totalProductos += pallets;
        }
      }
    }
  }

  // Calcular promedio ponderado de vida útil
  return totalProductos > 0 ? sumaPorcentajes / totalProductos : 0;
}

// Función para generar un gradiente ampliado verde-amarillo-rojo
function generarColorGradiente(porcentaje, sinDatos = false) {
  if (sinDatos) {
    return "#CCCCCC"; // Color gris claro para ubicaciones sin datos
  }

  // Limitar el porcentaje entre 0 y 100
  porcentaje = Math.max(0, Math.min(100, porcentaje));

  if (porcentaje <= 25) {
    return "#00FF00";
  } else if (porcentaje <= 50) {
    return "#FFFF00";
  } else if (porcentaje <= 75) {
    return "#FF8800"; // Naranjo
  } else {
    return "#FF0000";
  }
}

function generarColorGradienteVidaUtil(porcentaje) {
  // 100% vida útil (verde) -> 0% vida útil (rojo)
  porcentaje = typeof porcentaje === "number" ? porcentaje : 0;
  porcentaje = Math.max(0, Math.min(100, porcentaje));

  if (porcentaje >= 75) {
    // Verde (75-100%)
    return "#00FF00";
  } else if (porcentaje >= 50) {
    // Amarillo (50-75%)
    return "#FFFF00";
  } else if (porcentaje >= 25) {
    // Naranjo (25-50%)
    return "#FF8800"; // Naranjo
  } else {
    // Rojo (0-25%)
    return "#FF0000";
  }
}

// Función para obtener la URL del último SVG generado (útil para APIs o servicios web)
function limpiarPropiedadesEspecificas() {
  var documentProperties = PropertiesService.getDocumentProperties();

  // Eliminar propiedades específicas
  documentProperties.deleteProperty("ultimoSVG_ID");
  documentProperties.deleteProperty("file_name");
  documentProperties.deleteProperty("ultimoSVG_URL");

  Logger.log("Propiedades específicas eliminadas");
}

// Se llama en scripts.html
function actualizarTipoMapa(tipoMapa) {
  try {
    // Guardar la preferencia del usuario
    PropertiesService.getDocumentProperties().setProperty(
      "tipoMapaSeleccionado",
      tipoMapa
    );

    return {
      status: "success",
      message: "Tipo de mapa actualizado",
    };
  } catch (error) {
    Logger.log("Error al actualizar tipo de mapa: " + error.toString());
    return {
      status: "error",
      message: "Error: " + error.toString(),
    };
  }
}

function prepararDatosUbicaciones(ubicaciones) {
  const datosUbicaciones = {
    ocupacion: {},
    vencimiento: {},
  };

  // Procesar cada ubicación
  Object.entries(ubicaciones).forEach(([id, ubicacion]) => {
    const capacidad_maxima = ubicacion.capacidad_maxima || 0;
    const area = ubicacion.area || "";
    const utilizado = ubicacion.utilizado || 0;

    // Datos comunes
    const datosComunes = {
      area,
      capacidad_maxima,
      utilizado,
      niveles: ubicacion.niveles || {},
    };

    // Datos para ocupación
    const porcentajeOcupacion =
      capacidad_maxima > 0 ? (utilizado / capacidad_maxima) * 100 : 0;
    datosUbicaciones.ocupacion[id] = {
      ...datosComunes,
      porcentaje: porcentajeOcupacion.toFixed(1),
      color:
        (utilizado / capacidad_maxima) * 100 <= 0
          ? "#2D572C"
          : generarColorGradiente(porcentajeOcupacion),
    };

    // Datos para vencimiento
    const infoVidaUtil = obtenerCategoriasVidaUtil(ubicacion);
    const categoriasVidaUtil = infoVidaUtil.categorias;
    const conteoVidaUtil = infoVidaUtil.conteo;
    const porcentajeVidaUtil = calcularPorcentajeVidaUtil(ubicacion) || 0;

    datosUbicaciones.vencimiento[id] = {
      ...datosComunes,
      porcentaje: porcentajeVidaUtil.toFixed(1),
      color:
        utilizado <= 0
          ? "#666666"
          : generarColorGradienteVidaUtil(porcentajeVidaUtil),
      categoriasVidaUtil: categoriasVidaUtil,
      conteoVidaUtil: conteoVidaUtil,
      multipleVidasUtiles: categoriasVidaUtil.length > 1,
    };
  });

  return datosUbicaciones;
}

function generarMapaSKU(ubicaciones) {
  const mapaSKU = {};

  // Recorrer todas las ubicaciones
  Object.entries(ubicaciones).forEach(([id, ubicacion]) => {
    console.log("ubicacion:", ubicacion);
    if (ubicacion.niveles) {
      // Procesar cada nivel
      Object.entries(ubicacion.niveles).forEach(([nivel, infoNivel]) => {
        console.log("Info nivel:", infoNivel);
        // Nueva estructura con SKUs agrupados
        if (infoNivel.skus && typeof infoNivel.skus === "object") {
          Object.entries(infoNivel.skus).forEach(([codigoSku, skuInfo]) => {
            // Inicializar el SKU si no existe
            console.log("skuInfo", skuInfo);
            if (!mapaSKU[codigoSku]) {
              mapaSKU[codigoSku] = {
                descripcion: skuInfo.descripcion || "",
                ubicaciones: [],
              };
            }

            // Añadir esta ubicación a la lista
            mapaSKU[codigoSku].ubicaciones.push({
              ubicacion: extraerIDSimple(nivel),
              cantidad: skuInfo.pallets || 0,
            });
          });
        }
        // Compatibilidad con estructura antigua
        else if (
          infoNivel.sku &&
          typeof infoNivel.sku === "string" &&
          infoNivel.sku.trim() !== ""
        ) {
          const sku = infoNivel.sku;

          // Inicializar el SKU si no existe
          if (!mapaSKU[sku]) {
            mapaSKU[sku] = {
              descripcion: infoNivel.descripcion || "",
              ubicaciones: [],
            };
          }

          // Añadir esta ubicación a la lista
          mapaSKU[sku].ubicaciones.push({
            ubicacion: extraerIDSimple(nivel),
            cantidad: infoNivel.utilizado || 1,
          });
        }
      });
    }
  });

  return mapaSKU;
}
