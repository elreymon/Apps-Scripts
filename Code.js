// ===========================================================================
// Procesa la tabla T_INCOMING aplicando filtros y ordenación específicos
// v1.0
// ===========================================================================

// Variables globales para configuración
const SHEET_NAME = "INCOMING";
const TABLE_NAME = "T_INCOMING";

// Configuración de filtros
const FILTERS_CONFIG = {
  TITULO: {
    column: 'B', // TITULO
    excludeSubstrings: ["infantil", "mujer", "deport", "centros de mayores"],
    caseSensitive: false
  },
  DIAS_SEMANA: {
    column: 'F', // DIAS-SEMANA
    requireAny: ["V", "S", "D"]
  },
  GRATUITO: {
    column: 'D', // GRATUITO
    excludeValues: ["1"]
  },
  TIPO: {
    column: 'X', // TIPO
    excludeSubstrings: ["Flamenco", "CuentacuentosTiteresMarionetas"]
  },
  AUDIENCIA: {
    column: 'Y', // AUDIENCIA
    excludeSubstrings: ["Niños", "Familias", "Mayores"]
  },
  DISTRITO: {
    column: 'V', // DISTRITO-INSTALACION
    excludeValues: ["USERA", "VILLAVERDE"]
  }
};

// Configuración de ordenación
const SORT_CONFIG = {
  PRIMARY: {
    column: 'X', // TIPO
    priority: "ProgramacionDestacadaAgendaCultura"
  },
  SECONDARY: {
    column: 'F', // DIAS-SEMANA
    priority: "V"
  }
};

/**
 * Función principal requerida para aplicaciones web
 */
function doGet() {
  console.log("=== INICIANDO PROCESAMIENTO DE EVENTOS ===");
  
  try {
    const result = processEventsTable();
    
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'success',
        message: 'Procesamiento completado exitosamente',
        processedRows: result.processedRows,
        remainingRows: result.remainingRows
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    console.error("Error en doGet():", error.toString());
    
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error',
        message: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } // try
}

/**
 * Función principal que procesa la tabla de eventos
 * @return {Object} Resultado del procesamiento
 */
function processEventsTable() {
  console.log("Iniciando procesamiento de tabla T_INCOMING");
  
  // Obtener la hoja de cálculo y validar
  const sheet = getAndValidateSheet();
  if (!sheet) {
    throw new Error("No se pudo acceder a la hoja INCOMING");
  }
  
  // Obtener datos iniciales
  const initialData = getTableData(sheet);
  const initialRowCount = initialData.length;
  console.log(`Filas iniciales encontradas: ${initialRowCount}`);
  
  if (initialRowCount === 0) {
    console.log("No hay datos para procesar");
    return { processedRows: 0, remainingRows: 0 };
  }
  
  // Aplicar filtros
  console.log("=== APLICANDO FILTROS ===");
  let filteredData = applyAllFilters(initialData);
  const filteredRowCount = filteredData.length;
  console.log(`Filas después de filtrado: ${filteredRowCount} (eliminadas: ${initialRowCount - filteredRowCount})`);
  
  // Aplicar ordenación
  console.log("=== APLICANDO ORDENACIÓN ===");
  filteredData = applySorting(filteredData);
  console.log("Ordenación completada");
  
  // Escribir datos procesados de vuelta a la hoja
  updateSheetData(sheet, filteredData);
  
  console.log("=== PROCESAMIENTO COMPLETADO ===");
  return {
    processedRows: initialRowCount,
    remainingRows: filteredRowCount
  };
}

/**
 * Obtiene y valida la hoja de cálculo
 * @return {Sheet|null} La hoja validada o null si hay error
 */
function getAndValidateSheet() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    console.log(`Documento activo: ${spreadsheet.getName()}`);
    
    const sheet = spreadsheet.getSheetByName(SHEET_NAME);
    if (!sheet) {
      console.error(`Hoja '${SHEET_NAME}' no encontrada`);
      return null;
    }
    
    console.log(`Hoja '${SHEET_NAME}' encontrada exitosamente`);
    return sheet;
    
  } catch (error) {
    console.error("Error al acceder a la hoja:", error.toString());
    return null;
  }
}

/**
 * Obtiene todos los datos de la tabla
 * @param {Sheet} sheet - La hoja de cálculo
 * @return {Array} Array de arrays con los datos
 */
function getTableData(sheet) {
  console.log("Obteniendo datos de la tabla...");
  
  try {
    // Obtener todas las filas con datos (excluyendo encabezado)
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    
    if (lastRow <= 1) {
      console.log("No hay datos en la tabla (solo encabezado o vacía)");
      return [];
    }
    
    // Obtener datos desde fila 2 (saltando encabezado) hasta la última fila
    const range = sheet.getRange(2, 1, lastRow - 1, lastColumn);
    const data = range.getValues();
    
    console.log(`Datos obtenidos: ${data.length} filas x ${lastColumn} columnas`);
    return data;
    
  } catch (error) {
    console.error("Error al obtener datos:", error.toString());
    return [];
  }
}

/**
 * Aplica todos los filtros configurados
 * @param {Array} data - Datos originales
 * @return {Array} Datos filtrados
 */
function applyAllFilters(data) {
  let filteredData = [...data]; // Copia para no mutar el original
  
  // Aplicar filtro por TÍTULO
  filteredData = applyTitleFilter(filteredData);
  
  // Aplicar filtro por DÍAS-SEMANA
  filteredData = applyDaysFilter(filteredData);
  
  // Aplicar filtro por GRATUITO
  filteredData = applyFreeFilter(filteredData);
  
  // Aplicar filtro por TIPO
  filteredData = applyTypeFilter(filteredData);
  
  // Aplicar filtro por AUDIENCIA
  filteredData = applyAudienceFilter(filteredData);
  
  // Aplicar filtro por DISTRITO
  filteredData = applyDistrictFilter(filteredData);
  
  return filteredData;
}

/**
 * Aplica filtro por TÍTULO - elimina filas con subcadenas específicas
 * @param {Array} data - Datos a filtrar
 * @return {Array} Datos filtrados
 */
function applyTitleFilter(data) {
  const config = FILTERS_CONFIG.TITULO;
  const columnIndex = getColumnIndex(config.column);
  const initialCount = data.length;
  
  console.log(`Aplicando filtro TÍTULO (columna ${config.column})`);
  
  const filtered = data.filter(row => {
    const cellValue = String(row[columnIndex] || '');
    const valueToCheck = config.caseSensitive ? cellValue : cellValue.toLowerCase();
    
    const hasExcludedSubstring = config.excludeSubstrings.some(substring => {
      const substringToCheck = config.caseSensitive ? substring : substring.toLowerCase();
      return valueToCheck.includes(substringToCheck);
    });
    
    return !hasExcludedSubstring; // Mantener fila si NO contiene subcadenas excluidas
  });
  
  console.log(`Filtro TÍTULO: ${initialCount} -> ${filtered.length} (eliminadas: ${initialCount - filtered.length})`);
  return filtered;
}

/**
 * Aplica filtro por DÍAS-SEMANA - mantiene solo filas que contengan V, S o D
 * @param {Array} data - Datos a filtrar
 * @return {Array} Datos filtrados
 */
function applyDaysFilter(data) {
  const config = FILTERS_CONFIG.DIAS_SEMANA;
  const columnIndex = getColumnIndex(config.column);
  const initialCount = data.length;
  
  console.log(`Aplicando filtro DÍAS-SEMANA (columna ${config.column})`);
  
  const filtered = data.filter(row => {
    const cellValue = String(row[columnIndex] || '');
    
    const hasRequiredDay = config.requireAny.some(day => cellValue.includes(day));
    return hasRequiredDay; // Mantener fila si contiene alguno de los días requeridos
  });
  
  console.log(`Filtro DÍAS-SEMANA: ${initialCount} -> ${filtered.length} (eliminadas: ${initialCount - filtered.length})`);
  return filtered;
}

/**
 * Aplica filtro por GRATUITO - elimina filas que contengan "1"
 * @param {Array} data - Datos a filtrar
 * @return {Array} Datos filtrados
 */
function applyFreeFilter(data) {
  const config = FILTERS_CONFIG.GRATUITO;
  const columnIndex = getColumnIndex(config.column);
  const initialCount = data.length;
  
  console.log(`Aplicando filtro GRATUITO (columna ${config.column})`);
  
  const filtered = data.filter(row => {
    const cellValue = String(row[columnIndex] || '');
    
    const hasExcludedValue = config.excludeValues.some(value => cellValue === value);
    return !hasExcludedValue; // Mantener fila si NO contiene valores excluidos
  });
  
  console.log(`Filtro GRATUITO: ${initialCount} -> ${filtered.length} (eliminadas: ${initialCount - filtered.length})`);
  return filtered;
}

/**
 * Aplica filtro por TIPO - elimina filas con subcadenas específicas
 * @param {Array} data - Datos a filtrar
 * @return {Array} Datos filtrados
 */
function applyTypeFilter(data) {
  const config = FILTERS_CONFIG.TIPO;
  const columnIndex = getColumnIndex(config.column);
  const initialCount = data.length;
  
  console.log(`Aplicando filtro TIPO (columna ${config.column})`);
  
  const filtered = data.filter(row => {
    const cellValue = String(row[columnIndex] || '');
    
    const hasExcludedSubstring = config.excludeSubstrings.some(substring => {
      return cellValue.includes(substring);
    });
    
    return !hasExcludedSubstring; // Mantener fila si NO contiene subcadenas excluidas
  });
  
  console.log(`Filtro TIPO: ${initialCount} -> ${filtered.length} (eliminadas: ${initialCount - filtered.length})`);
  return filtered;
}

/**
 * Aplica filtro por AUDIENCIA - elimina filas con subcadenas específicas
 * @param {Array} data - Datos a filtrar
 * @return {Array} Datos filtrados
 */
function applyAudienceFilter(data) {
  const config = FILTERS_CONFIG.AUDIENCIA;
  const columnIndex = getColumnIndex(config.column);
  const initialCount = data.length;
  
  console.log(`Aplicando filtro AUDIENCIA (columna ${config.column})`);
  
  const filtered = data.filter(row => {
    const cellValue = String(row[columnIndex] || '');
    
    const hasExcludedSubstring = config.excludeSubstrings.some(substring => {
      return cellValue.includes(substring);
    });
    
    return !hasExcludedSubstring; // Mantener fila si NO contiene subcadenas excluidas
  });
  
  console.log(`Filtro AUDIENCIA: ${initialCount} -> ${filtered.length} (eliminadas: ${initialCount - filtered.length})`);
  return filtered;
}

/**
 * Aplica filtro por DISTRITO - elimina distritos específicos
 * @param {Array} data - Datos a filtrar
 * @return {Array} Datos filtrados
 */
function applyDistrictFilter(data) {
  const config = FILTERS_CONFIG.DISTRITO;
  const columnIndex = getColumnIndex(config.column);
  const initialCount = data.length;
  
  console.log(`Aplicando filtro DISTRITO (columna ${config.column})`);
  
  const filtered = data.filter(row => {
    const cellValue = String(row[columnIndex] || '');
    
    const hasExcludedValue = config.excludeValues.some(value => cellValue === value);
    return !hasExcludedValue; // Mantener fila si NO contiene valores excluidos
  });
  
  console.log(`Filtro DISTRITO: ${initialCount} -> ${filtered.length} (eliminadas: ${initialCount - filtered.length})`);
  return filtered;
}

/**
 * Aplica ordenación según criterios especificados
 * @param {Array} data - Datos a ordenar
 * @return {Array} Datos ordenados
 */
function applySorting(data) {
  const primaryConfig = SORT_CONFIG.PRIMARY;
  const secondaryConfig = SORT_CONFIG.SECONDARY;
  const primaryColumnIndex = getColumnIndex(primaryConfig.column);
  const secondaryColumnIndex = getColumnIndex(secondaryConfig.column);
  
  console.log(`Aplicando ordenación: Primario por ${primaryConfig.column}, Secundario por ${secondaryConfig.column}`);
  
  const sorted = [...data].sort((a, b) => {
    // Criterio primario: TIPO con "ProgramacionDestacadaAgendaCultura" primero
    const aPrimaryValue = String(a[primaryColumnIndex] || '');
    const bPrimaryValue = String(b[primaryColumnIndex] || '');
    const aPrimaryPriority = aPrimaryValue.includes(primaryConfig.priority);
    const bPrimaryPriority = bPrimaryValue.includes(primaryConfig.priority);
    
    if (aPrimaryPriority && !bPrimaryPriority) return -1; // a primero
    if (!aPrimaryPriority && bPrimaryPriority) return 1;  // b primero
    
    // Si ambos tienen la misma prioridad primaria, aplicar criterio secundario
    if (aPrimaryPriority === bPrimaryPriority) {
      const aSecondaryValue = String(a[secondaryColumnIndex] || '');
      const bSecondaryValue = String(b[secondaryColumnIndex] || '');
      const aSecondaryPriority = aSecondaryValue.includes(secondaryConfig.priority);
      const bSecondaryPriority = bSecondaryValue.includes(secondaryConfig.priority);
      
      if (aSecondaryPriority && !bSecondaryPriority) return -1; // a primero
      if (!aSecondaryPriority && bSecondaryPriority) return 1;  // b primero
    }
    
    return 0; // Mantener orden existente si no hay diferencias
  });
  
  console.log("Ordenación aplicada exitosamente");
  return sorted;
}

/**
 * Actualiza los datos en la hoja de cálculo
 * @param {Sheet} sheet - La hoja de cálculo
 * @param {Array} data - Los datos procesados
 */
function updateSheetData(sheet, data) {
  console.log("Actualizando datos en la hoja...");
  
  try {
    // Limpiar datos existentes (mantener encabezado)
    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();
    
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, lastColumn).clearContent();
      console.log("Datos anteriores eliminados");
    }
    
    // Escribir nuevos datos si existen
    if (data.length > 0) {
      const range = sheet.getRange(2, 1, data.length, data[0].length);
      range.setValues(data);
      console.log(`${data.length} filas escritas exitosamente`);
    } else {
      console.log("No hay datos para escribir");
    }
    
  } catch (error) {
    console.error("Error al actualizar datos:", error.toString());
    throw error;
  }
}

/**
 * Convierte letra de columna a índice (base 0)
 * @param {string} columnLetter - Letra de la columna (A, B, C, etc.)
 * @return {number} Índice de la columna
 */
function getColumnIndex(columnLetter) {
  let index = 0;
  for (let i = 0; i < columnLetter.length; i++) {
    index = index * 26 + (columnLetter.charCodeAt(i) - 64); // A=1, B=2, etc.
  }
  return index - 1; // Convertir a base 0
}

/**
 * Función de utilidad para testing/debugging manual
 */
function testProcessing() {
  console.log("=== PRUEBA MANUAL INICIADA ===");
  try {
    const result = processEventsTable();
    console.log("Resultado:", result);
  } catch (error) {
    console.error("Error en prueba:", error.toString());
  }
  console.log("=== PRUEBA MANUAL FINALIZADA ===");
}