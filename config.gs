/**
 * Configuración global de la aplicación
 */

// Configuración de la bodega MODELO (principal)
const CONFIG = {
  // Archivos y carpetas
  SVG_BASE: "mapa_bodega.svg",
  FOLDER_ID: "1Mt6cZimqlaGz59C5fLjKNgKSe1MRVPNe",
  SUBCARPETA_EXCEL_ID: "1-VnhciR3gt8f3IriYym7elruMyej1lLR",

  // Áreas de la bodega
  AREAS_CONSIDERADAS: [
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
  ],

  // Configuración de procesamiento
  PROCESSING: {
    BATCH_SIZE: 500,
    CACHE_DURATION: 1800, // 30 minutos
    CONVERSION_WAIT: 1500, // ms
    SHEET_WAIT: 2000, // ms - espera después de conversión de Excel
    RETRY_ATTEMPTS: 3,
  },

  // Configuración de colores para ocupación
  COLORS: {
    OCUPACION: {
      VACIO: "#2D572C",
      BAJO: "#00FF00", // 0-25%
      MEDIO: "#FFFF00", // 25-50%
      ALTO: "#FF8800", // 50-75%
      COMPLETO: "#FF0000", // 75-100%
    },
    VENCIMIENTO: {
      SIN_DATOS: "#666666",
      ALTO: "#00FF00", // 75-100%
      MEDIO: "#FFFF00", // 50-75%
      BAJO: "#FF8800", // 25-50%
      CRITICO: "#FF0000", // 0-25%
    },
    MOVIMIENTOS: {
      SIN_MOVIMIENTO: "#2D572C",
      BAJO: "#00FF00", // 0-25 percentil
      MEDIO: "#FFFF00", // 25-50 percentil
      ALTO: "#FF8800", // 50-75 percentil
      MUY_ALTO: "#FF0000", // 75-100 percentil
    },
  },

  // Configuración NPR (comentada para referencia)
  NPR: {
    // SVG_BASE: "VNA SEMI_layout.svg",
    // FOLDER_ID: "11EkkJ6G1plZEv8asYje5N5B2_XPyTeYa",
    // INVENTARIO_SHEET_ID: "1mozdeFrNIjUxOYxuafWDbqUdI2JW2HmlhxJt2xVOK5Y",
    // SUBCARPETA_EXCEL_ID: "1J4502AyItDtmL-yf2o9y2moP4_ajvNpy",
    // AREAS_CONSIDERADAS: ["VNA SEMI AUTOMATICO", "AREA ALMACENAMIENTO"]
  },
};

// Configuración de índices de columnas para diferentes fuentes de datos
const COLUMN_INDICES = {
  BLUEYONDER: {
    INVENTARIO: [
      { key: "area", name: "Área" },
      { key: "zonaTrabajo", name: "Zona de trabajo" },
      { key: "ubicacion", name: "Ubicación" },
      { key: "lpn", name: "LPN" },
      { key: "sku", name: "Artículo" },
      { key: "descripcion", name: "Descripcion" },
      { key: "cajas", name: "Cantidad" },
      { key: "fifo", name: "Fecha FIFO" },
      { key: "fabricacion", name: "Fecha de fabricacion" },
      { key: "caducidad", name: "Fecha de caducidad" },
      { key: "recepcion", name: "Fecha de recepcion" },
      { key: "perfilAntiguedad", name: "Nombre de perfil de antiguedad" },
      { key: "diasHastaCaducidad", name: "Dias hasta su caducidad" },
      { key: "ultimoMov", name: "Ultima fecha de movimiento" },
    ],
    UBICACIONES: [
      { key: "ubicacion", name: "Ubicación" },
      { key: "area", name: "Área" },
    ],
  },

  WEP: {
    INVENTARIO: [
      { key: "ubicacion", name: "Localidad" },
      { key: "sku", name: "SKU" },
      { key: "descripcion", name: "Producto" },
      { key: "lpn", name: "LPN" },
      { key: "cajas", name: "Cantidad" },
      { key: "caducidad", name: "Caducidad" },
    ],
    UBICACIONES: [
      { key: "ubicacion", name: "Localidad" },
      { key: "area", name: "Área" },
      { key: "capacidad_maxima", name: "Cap. Máxima" },
    ],
  },

  CAJAS_PALLET: [
    { key: "sku", name: "Artículo" },
    { key: "cajas_x_pallet", name: "Cajas X Pallets" },
    { key: "vida_util", name: "SKU_VIDA_UTIL" },
  ],
};

// Funciones auxiliares para obtener configuración
function getConfig() {
  return CONFIG;
}

function getColumnIndices(source, type) {
  const sourceUpper = source.toUpperCase();
  const typeUpper = type.toUpperCase();

  if (COLUMN_INDICES[sourceUpper] && COLUMN_INDICES[sourceUpper][typeUpper]) {
    return COLUMN_INDICES[sourceUpper][typeUpper];
  }

  return COLUMN_INDICES.CAJAS_PALLET;
}

function getProcessingConfig() {
  return CONFIG.PROCESSING;
}

function getColors() {
  return CONFIG.COLORS;
}
