/**
 * 🧠 MOTOR FINANCIERO CENTRAL (Versión Modular)
 */
import * as XLSX from 'xlsx';

export function normalizeText(text) {
    if (!text) return "";
    // Normalizar: minúsculas, sin acentos y remover puntuación común de cabeceras/paréntesis
    return text.toString()
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/[.:;()]/g, " ")
        .replace(/\s+/g, " ")
        .trim();
}

export function cleanNumber(val) {
    if (typeof val === 'number') return val;
    if (val === null || val === undefined) return 0;
    
    let cleaned = val.toString().trim().replace(/\u00A0/g, ' '); // Handle non-breaking spaces
    if (!cleaned) return 0;

    // Si no es un número puro, intentamos extraer la parte numérica (ej: "Tasa 58.50" -> 58.50)
    if (isNaN(cleaned.replace(/,/g, ''))) {
        const match = cleaned.match(/-?[\d,.]+/);
        if (match) cleaned = match[0];
    }

    // Handle (1,234.56) notation for negative numbers
    let isNegative = false;
    if (cleaned.startsWith('(') && cleaned.endsWith(')')) {
        isNegative = true;
        cleaned = cleaned.substring(1, cleaned.length - 1);
    } else if (cleaned.startsWith('-')) {
        isNegative = true;
        cleaned = cleaned.substring(1);
    }

    // Advanced thousands separator handling
    if (cleaned.includes(',') && cleaned.includes('.')) {
        if (cleaned.lastIndexOf(',') > cleaned.lastIndexOf('.')) {
            cleaned = cleaned.replace(/\./g, '').replace(',', '.');
        } else {
            cleaned = cleaned.replace(/,/g, '');
        }
    } else if (cleaned.includes(',')) {
        const parts = cleaned.split(',');
        if (parts[parts.length - 1].length === 3) {
            cleaned = cleaned.replace(/,/g, '');
        } else {
            cleaned = cleaned.replace(',', '.');
        }
    }

    cleaned = cleaned.replace(/[$\s%]/g, '');
    let num = parseFloat(cleaned);
    if (isNaN(num)) return 0;
    return isNegative ? -num : num;
}

export function formatCurrency(val) {
    if (val === null || val === undefined) return "$0.0M";
    const absVal = Math.abs(val);
    
    // Si el valor viene "crudo" en unidades regulares (ej. > 200,000), lo dividimos por 1 millón
    if (absVal >= 200000) {
        const millions = absVal / 1000000;
        const formatted = new Intl.NumberFormat('en-US', { 
            minimumFractionDigits: 1, 
            maximumFractionDigits: 1 
        }).format(millions);
        return `${val < 0 ? '-' : ''}$${formatted}M`;
    } else {
        // El archivo fuente ya tiene los datos expresados en millones (ej. 5.4)
        // Formateamos "tal cual" sin hacer división, asumiendo su escala natural.
        const formatted = new Intl.NumberFormat('en-US', { 
            minimumFractionDigits: 1, 
            maximumFractionDigits: 1 
        }).format(absVal);
        return `${val < 0 ? '-' : ''}$${formatted}M`;
    }
}

export function formatRawCurrency(val) {
    if (val === null || val === undefined) return "$0.0";
    const absVal = Math.abs(val);
    
    // Convert to millions for congruency
    const millions = absVal / 1000000;
    
    const formatted = new Intl.NumberFormat('en-US', { 
        minimumFractionDigits: 1, 
        maximumFractionDigits: 1 
    }).format(millions);
    return `${val < 0 ? '-' : ''}$${formatted}M`;
}

export function formatPercent(val) { 
    return (val * 100).toFixed(1) + '%'; 
}

export function financialEngine(workbook) {
    let sheets = {};
    workbook.SheetNames.forEach(name => {
        const normName = normalizeText(name);
        sheets[normName] = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1 });
    });

    const sheetKeys = Object.keys(sheets);
    
    // Prioritize "PA" or "Seguimiento" as per user context
    const pnlKey = sheetKeys.find(s => (/pa|seguimiento|gerencial|p&l|resultado|income|ganancia/i.test(s))) || sheetKeys[0];
    const balanceKey = sheetKeys.find(s => s.includes("balance sheet mdop")) || 
                       sheetKeys.find(s => /balance|situacion|estado/i.test(s) && !/p&l|resultado/i.test(s));
    const cashflowKey = sheetKeys.find(s => /cash|flujo/i.test(s));

    if (pnlKey && sheets[pnlKey]) {
        const result = processFinancialStatements(sheets, pnlKey, balanceKey, cashflowKey);
        if (!result.error && result.data && result.data.length > 0) {
            result.modelType = "Reporte PA / Estados Financieros";
            return result;
        }
    }

    const tbKey = sheetKeys.find(k => k === "tb" || k.includes("trial balance") || k.includes("balanza") || k.includes("data"));
    const setupKey = sheetKeys.find(k => k === "setup" || k.includes("mapeo") || k.includes("config") || k.includes("mapping"));

    function isRealTB(sheet) {
        if (!sheet || sheet.length < 5) return false;
        let numericAccountCount = 0;
        const sample = sheet.slice(0, 50);
        sample.forEach(row => {
            if (!row) return;
            row.forEach(cell => {
                const val = String(cell).trim();
                if (/^\d{4,}/.test(val)) numericAccountCount++;
            });
        });
        return numericAccountCount > 3;
    }

    const tbSheet = tbKey ? sheets[tbKey] : null;
    const isTB = isRealTB(tbSheet);

    if (isTB && tbKey && setupKey) {
        const result = processTBSetup(sheets, tbKey, setupKey);
        if (!result.error) {
            result.modelType = "TB + Setup (Contable)";
            return result;
        }
    }

    const resultWide = processWide(sheets);
    if (!resultWide.error) {
        resultWide.modelType = "Wide Format (Reporte Gerencial)";
        return resultWide;
    }

    return { error: `No se detectó un modelo válido. El archivo debe contener una hoja llamada 'P&L', 'Balance' o 'TB'. Hojas encontradas: ${sheetKeys.join(", ")}` };
}

// ... (Porting all functions from index.html to here)
// I will include the full logic to ensure it works as before

function findRowByKeywords(rows, keywords, targetRowIdxHint = null) {
    let bestRow = null;
    let maxScore = -1;

    rows.forEach((row, idx) => {
        if (!row || row.length < 2) return;
        // Revisar más columnas (hasta la 10) por si el label está desplazado
        for (let i = 0; i < Math.min(row.length, 10); i++) {
            const cell = row[i];
            if (!cell) continue;
            const label = normalizeText(cell);
            
            if (keywords.some(k => label === k || (k.length > 3 && label.includes(k)))) {
                let numCount = 0;
                let potentialTotal = 0;
                for (let j = 1; j < row.length; j++) {
                    const val = cleanNumber(row[j]);
                    if (val !== 0) {
                        numCount++;
                        potentialTotal = Math.max(potentialTotal, Math.abs(val));
                    }
                }

                let score = numCount;
                // Prioridad alta a coincidencias exactas con keywords importantes
                if (keywords.some(k => label === k)) score += 30;
                
                if (label.includes("total") || label.includes("sum") || label.includes("consolidado")) score += 15;
                if (label.includes("neto") || label.includes("final") || label.includes("ejercicio")) score += 20;

                // Si el usuario nos dio una pista de fila (ej: fila 61 en excel es idx 60)
                if (targetRowIdxHint !== null) {
                    if (Math.abs(idx - targetRowIdxHint) <= 5) score += 50; // Gran bono si está cerca de la fila 61
                }
                
                if (label.includes("%") || label.includes("var") || label.includes("crecimiento")) score -= 15;
                
                if (score > maxScore) {
                    maxScore = score;
                    bestRow = row;
                }
                break; 
            }
        }
    });
    return bestRow;
}

function detectSegments(rows, segmentKeywords) {
    const segments = {};
    rows.forEach(row => {
        if (!row) return;
        for (let i = 0; i < Math.min(row.length, 5); i++) {
            const cell = row[i];
            if (!cell) continue;
            const label = normalizeText(cell);
            
            segmentKeywords.forEach(seg => {
                const normSeg = seg.toLowerCase();
                const regex = new RegExp(`\\b${normSeg}\\b`, 'i');
                if (regex.test(label)) {
                    const finalSegName = (seg === "P6" || seg === "BON") ? "BON" : seg;
                    
                    if (!segments[finalSegName]) segments[finalSegName] = { ventasRows: [], costosRows: [] };
                    
                    const hasNumbers = row.some((c, idx) => idx > i && cleanNumber(c) !== 0);
                    if (hasNumbers) {
                        if (label.includes("costo") || label.includes("costos")) {
                            segments[finalSegName].costosRows.push(row);
                        } else if (label.includes("venta") || label.includes("ingreso") || !label.includes("total")) {
                            segments[finalSegName].ventasRows.push(row);
                        }
                    }
                }
            });
        }
    });
    return segments;
}

export const FINANCIAL_KEYWORDS = {
    ingresos: ["ventas", "ingresos", "revenue", "ventas netas", "total ingresos", "facturacion", "servicios", "productos", "ventas totales"],
    costos: ["costo de ventas", "costos directos", "cogs", "cost of sales", "total costos", "costos de operacion"],
    opex: ["gastos operativos", "opex", "gastos de administracion", "total gastos operativos", "gastos de venta", "otros gastos operativos", "ggadm", "gastos generales", "total gastos", "operativos"],
    ebitda: ["ebitda", "utilidad operativa", "operating income", "uafida", "utilidad antes de", "resultado operativo", "margen operativo", "utilidad de operacion"],
    utilidad: ["utilidad neta", "net income", "resultado del ejercicio", "utilidad perdida", "beneficio neto", "resultado neto", "utilidad del periodo", "ganancia neta", "ganancia del ejercicio", "utilidad neta ejercicio", "utilidad neta periodo", "ganancia perdida ejercicio", "resultado del periodo", "resultado"],
    cashflow: ["cash flow", "flujo de caja", "flujo neto", "disponibilidad", "caja final", "efectivo", "flujo de efectivo"],
    tasa_cambio: ["tasa de cambio", "fx rate", "tipo de cambio", "tasa bpd", "tasa promedio", "t.c", "tc", "tasa", "cambio", "exchange"],
    // Nuevas Keywords para Hoja de Cash Flow
    cf_beginning: ["beginning cash balance", "efectivo inicial", "saldo inicial de efectivo", "caja inicial"],
    cf_operating: ["operating activities", "flujo de actividades de operacion", "actividades de operacion", "flujo de caja operativo"],
    cf_wc: ["change in working capital", "cambios en capital de trabajo", "variacion capital de trabajo", "working capital requirements"],
    cf_cxc: ["aumento)/disminucion en cuentas por cobrar", "cuentas por cobrar", "cxc", "accounts receivable"],
    cf_inv: ["aumento)/disminucion en inventario", "inventario", "inventarios", "inventory"],
    cf_cxp: ["aumento/(disminucion) en cuentas por pagar", "cuentas por pagar", "cxp", "accounts payable"],
    cf_capex: ["capex", "inversiones de capital", "desembolsos de capital", "adquisicion de activos", "capital expenditures"],
    cf_financing: ["financing activities", "flujo de actividades de financiamiento", "actividades de financiamiento"],
    cf_net_debt: ["aumento deuda neta", "variacion de deuda", "financiamiento neto", "deuda bancaria", "net debt", "repayment of debt"],
    cf_change: ["change in cash", "cambio en efectivo", "variacion neta de efectivo"],
    cf_ending: ["ending cash balance", "efectivo final", "saldo final de efectivo", "caja final"],
    cf_below_ebitda: ["below ebitda"],
    cf_taxes: ["taxes", "impuestos", "pago impuestos", "income taxes"],
    cf_dividends: ["dividends", "dividendos", "shareholders activities", "accionistas"],
    cf_interest: ["gastos de interes", "intereses", "interest expense", "financial expenses", "interests earned"],
    cf_extraordinary: ["gastos extraordinarios", "ingresos extraordinarios", "extraordinarios", "extraordinary items"],
    cf_dso: ["dso"],
    cf_dpo: ["dpo"],
    cf_dio: ["dio"]
};

function processFinancialStatements(sheets, pnlKey, balanceKey, cashflowKey) {
    const pnlSheet = sheets[pnlKey];
    const balanceSheet = balanceKey ? sheets[balanceKey] : null;

    const cashflowSheet = cashflowKey ? sheets[cashflowKey] : null;

    // Detectar si el Balance o P&L están en millones (mDOP)
    let isBalanceInMillions = (balanceKey && (normalizeText(balanceKey).includes("mdop") || normalizeText(balanceKey).includes("millones") || normalizeText(balanceKey).includes("mrd$"))) ||
                              (pnlKey && (normalizeText(pnlKey).includes("mdop") || normalizeText(pnlKey).includes("millones") || normalizeText(pnlKey).includes("mrd$"))) ||
                              (cashflowKey && (normalizeText(cashflowKey).includes("mdop") || normalizeText(cashflowKey).includes("millones") || normalizeText(cashflowKey).includes("mrd$")));
    
    // If no scale is detected by text, check values
    if (!isBalanceInMillions) {
        const detectScale = (sheet) => {
            if (!sheet) return false;
            let foundText = false;
            let smallValuesCount = 0;
            let nonZeroCount = 0;

            // Revisamos hasta 200 filas para estar seguros de capturar todo el contexto
            for (let i = 0; i < Math.min(sheet.length, 200); i++) {
                if (sheet[i]) {
                    const rowStr = normalizeText(sheet[i].join(" "));
                    if (rowStr.includes("mdop") || rowStr.includes("millones") || rowStr.includes("mrd$") || rowStr.includes("cifras en")) foundText = true;
                    
                    sheet[i].forEach(cell => {
                        const n = cleanNumber(cell);
                        if (n !== 0 && !isNaN(n)) {
                            nonZeroCount++;
                            if (Math.abs(n) < 1000000) smallValuesCount++;
                        }
                    });
                }
            }
            return foundText || (nonZeroCount > 5 && (smallValuesCount / nonZeroCount) > 0.7);
        };
        isBalanceInMillions = detectScale(pnlSheet) || detectScale(balanceSheet) || detectScale(cashflowSheet);
    }

    const getVal = (row, idx, scale = 1) => {
        if (!row || idx === undefined || idx === null) return 0;
        let val = cleanNumber(row[idx]);
        // Fallback: buscar en un rango de +/- 2 columnas (Excel suele tener variaciones de alineación)
        if (val === 0) {
            for (let offset of [1, -1, 2, -2]) {
                const checkVal = cleanNumber(row[idx + offset]);
                if (checkVal !== 0) {
                    val = checkVal;
                    break;
                }
            }
        }
        return val * scale;
    };

    const getBalanceVal = (row, idx) => {
        if (!row) return 0;
        let val = cleanNumber(row[idx]);
        
        // Fallback: buscar en un rango de +/- 2 columnas
        if (val === 0) {
            for (let offset of [1, -1, 2, -2]) {
                const checkVal = cleanNumber(row[idx + offset]);
                if (checkVal !== 0) {
                    val = checkVal;
                    break;
                }
            }
        }

        const concept = row[0] ? normalizeText(String(row[0])) : "";
        // Detectar si es un ratio (unitless) o moneda
        const isRatio = (concept.includes("ratio") || concept.includes("indice") || concept.includes("razon") ||
                         concept.includes("apalancamiento") || concept.includes("capacidad") || 
                         concept.includes("covenant") || concept.includes("corriente") ||
                         concept.includes("deuda neta/ebitda") || concept.includes(" x ") || concept.endsWith(" x")) && 
                         !concept.includes("cxp") && !concept.includes("cxc") && !concept.includes("pagar") && !concept.includes("cobrar");
        
        if (isRatio) return val; // No escalar ratios
        return val * (isBalanceInMillions ? 1000000 : 1);
    };

    const detailedOpexKeywords = {
        admin: ["gastos administrativos", "gastos de administracion", "administracion"],
        mercadeo: ["gastos de mercadeo", "mercadeo", "publicidad", "marketing"],
        comercial: ["gastos de ventas (comercial)", "gastos de ventas", "comercial", "gastos comerciales"],
        logistica: ["gastos de logistica", "logistica", "gastos logisticos"]
    };

    const pnlRows = {
        ingresos: findRowByKeywords(pnlSheet, FINANCIAL_KEYWORDS.ingresos),
        costos: findRowByKeywords(pnlSheet, FINANCIAL_KEYWORDS.costos),
        opex: findRowByKeywords(pnlSheet, FINANCIAL_KEYWORDS.opex),
        ebitda: findRowByKeywords(pnlSheet, FINANCIAL_KEYWORDS.ebitda),
        utilidad: findRowByKeywords(pnlSheet, FINANCIAL_KEYWORDS.utilidad, 60), // Hint: Fila 61 (index 60)
        cashflow: findRowByKeywords(pnlSheet, FINANCIAL_KEYWORDS.cashflow),
        tasa_cambio: findRowByKeywords(pnlSheet, FINANCIAL_KEYWORDS.tasa_cambio),
        // Detalle de OPEX
        admin: findRowByKeywords(pnlSheet, detailedOpexKeywords.admin),
        mercadeo: findRowByKeywords(pnlSheet, detailedOpexKeywords.mercadeo),
        comercial: findRowByKeywords(pnlSheet, detailedOpexKeywords.comercial),
        logistica: findRowByKeywords(pnlSheet, detailedOpexKeywords.logistica),
        tasa_cambio: findRowByKeywords(pnlSheet, FINANCIAL_KEYWORDS.tasa_cambio)
    };

    const segmentKeywords = ["BT5", "EVP", "BON", "P6"];
    const segmentRows = detectSegments(pnlSheet, segmentKeywords);

    if (!pnlRows.ingresos) return { error: "No se encontró la fila de 'Ingresos' en el P&L. Verifique que los nombres de las filas sean claros (ej: 'Ventas' o 'Ingresos')." };

    const balanceKeywords = {
        activos: ["total activos", "activos", "total activo", "activo total", "total de activos", "activos totales"],
        pasivos: ["total pasivos", "pasivos", "total pasivo", "pasivo total", "sumas iguales pasivo", "pasivos totales"],
        patrimonio: ["total patrimonio", "patrimonio", "capital", "total capital", "capital contable", "patrimonio neto", "total pasivo y patrimonio", "total pasivo y capital"]
    };

    const balanceRows = {
        activos: (balanceSheet ? findRowByKeywords(balanceSheet, balanceKeywords.activos) : null) || findRowByKeywords(pnlSheet, balanceKeywords.activos),
        pasivos: (balanceSheet ? findRowByKeywords(balanceSheet, balanceKeywords.pasivos) : null) || findRowByKeywords(pnlSheet, balanceKeywords.pasivos),
        patrimonio: (balanceSheet ? findRowByKeywords(balanceSheet, balanceKeywords.patrimonio) : null) || findRowByKeywords(pnlSheet, balanceKeywords.patrimonio),
        // Cuentas específicas para cálculo de Beneficio Neto si viene en 0
        gananciaAcumulada: (balanceSheet ? findRowByKeywords(balanceSheet, ["ganancia acumulada", "utilidad acumulada", "ganancias acumuladas", "utilidades acumuladas"]) : null),
        utilidadesRetenidas: (balanceSheet ? findRowByKeywords(balanceSheet, ["utilidades retenidas", "utilidad retenida"]) : null)
    };

    const cfRows = {
        beginning: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_beginning) : null,
        operating: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_operating) : null,
        wc: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_wc) : null,
        cxc: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_cxc) : null,
        inv: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_inv) : null,
        cxp: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_cxp) : null,
        capex: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_capex) : null,
        financing: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_financing) : null,
        netDebt: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_net_debt) : null,
        belowEbitda: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_below_ebitda) : null,
        taxes: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_taxes) : null,
        dividends: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_dividends) : null,
        interest: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_interest) : null,
        extraordinary: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_extraordinary) : null,
        change: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_change) : null,
        ending: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_ending) : null,
        dso: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_dso) : null,
        dpo: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_dpo) : null,
        dio: cashflowSheet ? findRowByKeywords(cashflowSheet, FINANCIAL_KEYWORDS.cf_dio) : null
    };

    // Helper to find data column indices for a given sheet based on target dates
    const findSheetIndices = (sheet) => {
        const indices = {};
        if (!sheet) return indices;
        
        const monthNames = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
        const shortMonths = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];

        for (let i = 0; i < Math.min(sheet.length, 50); i++) {
            const row = sheet[i];
            if (!row) continue;
            row.forEach((cell, j) => {
                let dateObj = null;
                if (cell instanceof Date) dateObj = cell;
                else if (typeof cell === 'number' && cell > 40000 && cell < 60000) dateObj = new Date((cell - 25569) * 86400 * 1000);
                else if (typeof cell === 'string') {
                    const val = normalizeText(cell);
                    const mIdx = monthNames.findIndex(m => val.includes(m));
                    const sIdx = shortMonths.findIndex(s => val.includes(s));
                    const finalMIdx = mIdx !== -1 ? mIdx : sIdx;
                    
                    if (finalMIdx !== -1) {
                        dateObj = new Date();
                        dateObj.setMonth(finalMIdx);
                        const yearMatch = val.match(/\d{2,4}/);
                        if (yearMatch) {
                            let y = parseInt(yearMatch[0]);
                            if (y < 100) y += 2000;
                            dateObj.setFullYear(y);
                        } else {
                            // Look for year in neighboring cells if not in string
                            for (let neighborIdx = Math.max(0, i-2); neighborIdx <= Math.min(sheet.length-1, i+2); neighborIdx++) {
                                const neighborRow = sheet[neighborIdx];
                                if (!neighborRow) continue;
                                const yearInRow = neighborRow.find(c => typeof c === 'number' && c >= 2020 && c <= 2026);
                                if (yearInRow) { dateObj.setFullYear(yearInRow); break; }
                            }
                        }
                    } else if (val.match(/20\d{2}/) && val.match(/\d{1,2}/)) {
                        const dateMatch = val.match(/(20\d{2})[-/](\d{1,2})/);
                        if (dateMatch) dateObj = new Date(parseInt(dateMatch[1]), parseInt(dateMatch[2]) - 1, 1);
                    }
                }

                if (dateObj) {
                    const y = dateObj.getFullYear();
                    if (y >= 2020 && y <= 2026) {
                        const dateKey = `${dateObj.getMonth()}-${y}`;
                        if (!indices[dateKey]) indices[dateKey] = j;
                    }
                }
            });
        }
        return indices;
    };

    const pnlIndices = findSheetIndices(pnlSheet);
    const balanceIndices = balanceSheet ? findSheetIndices(balanceSheet) : {};
    const cfIndices = cashflowSheet ? findSheetIndices(cashflowSheet) : {};
    
    // Unificar todas las fechas detectadas en ambos reportes
    const allDateKeys = new Set([...Object.keys(pnlIndices), ...Object.keys(balanceIndices), ...Object.keys(cfIndices)]);
    
    let dataPeriods = [];
    allDateKeys.forEach(key => {
        const [m, y] = key.split('-').map(Number);
        const d = new Date(y, m, 1);
        
        // 🚨 Filtro de seguridad: Solo permitir periodos hasta 2026 (pedido por usuario)
        // Y evitar fechas absurdamente lejanas en el pasado o futuro
        if (y >= 2020 && y <= 2026) {
            const pnlIdx = pnlIndices[key] !== undefined ? pnlIndices[key] : -1;
            const balanceIdx = balanceIndices[key] !== undefined ? balanceIndices[key] : -1;
            const cfIdx = cfIndices[key] !== undefined ? cfIndices[key] : -1;
            dataPeriods.push({ date: d, pnlIdx, balanceIdx, cfIdx });
        }
    });
    
    // Ordenar cronológicamente
    dataPeriods.sort((a, b) => a.date - b.date);

    if (dataPeriods.length === 0) {
        pnlRows.ingresos.forEach((cell, j) => {
            if (j === 0) return;
            const val = cleanNumber(cell);
            if (val !== 0 && !isNaN(val)) {
                const d = new Date();
                d.setMonth(d.getMonth() - (pnlRows.ingresos.length - j));
                dataPeriods.push({ date: d, pnlIdx: j, balanceIdx: j });
            }
        });
    }

    if (dataPeriods.length === 0) {
        return { error: "No se encontraron periodos o fechas válidas en las cabeceras." };
    }

    const getBalanceIdx = (date, pnlIdx) => {
        const key = `${date.getMonth()}-${date.getFullYear()}`;
        return balanceIndices[key] !== undefined ? balanceIndices[key] : pnlIdx;
    };

    const result = dataPeriods.map(point => {
        const pIdx = point.pnlIdx;
        const bIdx = point.balanceIdx;
        const cfIdx = point.cfIdx;

        const ingresos = pIdx !== -1 ? getVal(pnlRows.ingresos, pIdx) : 0;
        const costos = pIdx !== -1 ? getVal(pnlRows.costos, pIdx) : 0;
        const ebitda = pIdx !== -1 ? getVal(pnlRows.ebitda, pIdx) : 0;
        
        let opex = pIdx !== -1 ? getVal(pnlRows.opex, pIdx) : 0;
        if (ebitda !== 0 && opex === 0) {
            opex = Math.abs(ingresos - costos - ebitda);
        } else if (ebitda !== 0 && pIdx !== -1 && Math.abs(opex - (ingresos - costos - ebitda)) > 100) {
            const impliedOpex = Math.abs(ingresos - costos - ebitda);
            if (opex < impliedOpex * 0.5) opex = impliedOpex;
        }

        let utilidad = pIdx !== -1 ? getVal(pnlRows.utilidad, pIdx) : 0;
        
        // 🚨 CRITICAL FIX: Si la utilidad es 0, intentar calcularla por diferencia en el Balance
        if (utilidad === 0 && bIdx !== -1) {
            const gananciaAcum = getBalanceVal(balanceRows.gananciaAcumulada, bIdx);
            const utilRetenidas = getBalanceVal(balanceRows.utilidadesRetenidas, bIdx);
            if (gananciaAcum !== 0 || utilRetenidas !== 0) {
                utilidad = utilRetenidas - gananciaAcum;
            }
        }

        let cashflowVal = (pIdx !== -1 ? getVal(pnlRows.cashflow, pIdx) : 0) || (cfIdx !== -1 ? getVal(cfRows.change, cfIdx) : utilidad);

        const activos = bIdx !== -1 ? getBalanceVal(balanceRows.activos, bIdx) : 0;
        const pasivos = bIdx !== -1 ? getBalanceVal(balanceRows.pasivos, bIdx) : 0;
        const patrimonio = bIdx !== -1 ? getBalanceVal(balanceRows.patrimonio, bIdx) : 0;
        const tasaCambio = pIdx !== -1 ? getVal(pnlRows.tasa_cambio, pIdx) : 1;

        // Extraer Detalle de Cash Flow completo si existe
        const cashflowDetail = {};
        if (cashflowSheet && cfIdx !== -1) {
            Object.keys(cfRows).forEach(key => {
                const row = cfRows[key];
                if (row) cashflowDetail[key] = getVal(row, cfIdx);
            });
        }

        // Si el cambio neto viene en 0, intentar calcularlo por la suma de actividades
        if (cashflowVal === 0 && cashflowSheet && cfIdx !== -1) {
            const calculatedChange = (cashflowDetail.operating || 0) + (cashflowDetail.financing || 0) + (cashflowDetail.capex || 0);
            if (calculatedChange !== 0) cashflowVal = calculatedChange;
        }

        const segments = {};
        Object.entries(segmentRows).forEach(([name, data]) => {
            const sumRows = (rowList) => rowList.reduce((acc, row) => acc + (pIdx !== -1 ? getVal(row, pIdx) : 0), 0);
            segments[name] = {
                ventas: sumRows(data.ventasRows),
                costos: sumRows(data.costosRows)
            };
        });

        const fullRows = pnlSheet.filter(row => {
            if (!row || !row[0]) return false;
            const concept = normalizeText(row[0]);
            if (concept.includes("formatcode") || concept.includes("unnamed") || concept.length < 2) return false;
            return dataPeriods.some(p => p.pnlIdx !== -1 && (typeof row[p.pnlIdx] === 'number' || !isNaN(cleanNumber(row[p.pnlIdx]))));
        }).map(row => {
            const rowValues = {};
            dataPeriods.forEach(p => {
                rowValues[p.date.toLocaleDateString('es-ES', { month: 'short', year: 'numeric' })] = p.pnlIdx !== -1 ? getVal(row, p.pnlIdx) : 0;
            });
            const rawConcept = String(row[0]);
            let renamedConcept = (normalizeText(rawConcept) === "ventas p6") ? "Ventas BON" : rawConcept;
            if (normalizeText(renamedConcept) === "ganancia del periodo") renamedConcept = "Beneficio Neto del Periodo";
            return { concept: renamedConcept, values: rowValues };
        });

        const bSheetToUse = balanceSheet || pnlSheet;
        const balanceFullRows = bSheetToUse.filter(row => {
            if (!row || !row[0]) return false;
            const conceptStr = String(row[0]);
            const concept = normalizeText(conceptStr);
            if (concept.includes("formatcode") || concept.includes("unnamed") || concept.length < 2) return false;
            
            const isHeader = concept === "activos" || concept === "pasivos" || concept === "patrimonio" || 
                             concept === "capital" || concept === "pasivo y capital" || 
                             concept === "activo" || concept === "pasivo" ||
                             concept === "ingresos" || concept === "costos" || concept === "gastos";
            
            const isAccountingRule = concept.includes("ganancia acumulada") || concept.includes("utilidad acumulada") || 
                                    concept.includes("utilidades retenidas") || concept.includes("ganancia retenida") ||
                                    concept.includes("beneficio neto") || concept.includes("utilidad del ejercicio");

            if (isHeader && !isAccountingRule && !concept.includes("total")) return false;
            if (!isAccountingRule && (concept.includes("en mdop") || concept.includes("estado de situacion") || concept.includes("reporte pa"))) return false;
            
            const monthNamesArr = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
            const shortMonths = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
            if (monthNamesArr.some(m => concept.includes(m)) || shortMonths.some(s => concept.includes(s))) return false;
            
            const isTypicalBalance = concept.includes("activo") || concept.includes("pasivo") || 
                                     concept.includes("patrimonio") || concept.includes("efectivo") || 
                                     concept.includes("bancos") || concept.includes("cobrar") || 
                                     concept.includes("inventario") || concept.includes("propiedad") || 
                                     concept.includes("ppe") || concept.includes("prestamos") || 
                                     concept.includes("capital") || concept.includes("reserva") ||
                                     concept.includes("covenant") || concept.includes("deuda neta") ||
                                     concept.includes("ltm ebitda") || concept.includes("ebitda r12") || 
                                     concept.includes("deuda bruta") || concept.includes("deuda total") ||
                                     concept.includes("deuda subordinada") || concept.includes("deuda sin subordinada") ||
                                     concept.includes("apalancamiento") ||
                                     concept.includes("capacidad de pago") || concept.includes("capacidad") || 
                                     concept.includes("razon corriente") ||
                                     concept.includes("ganancia") || concept.includes("beneficio");

            const isNetIncomeInBalance = (concept.includes("utilidad") || concept.includes("ganancia") || concept.includes("beneficio") || concept.includes("ganancia")) && 
                                         (concept.includes("ejercicio") || concept.includes("periodo") || concept.includes("neto"));

            if (bSheetToUse === pnlSheet && !isTypicalBalance && !isNetIncomeInBalance) {
                const pnlStrict = ["ingresos", "ventas netas", "costo de ventas", "utilidad bruta", "ebitda", "ggadm", "ebit"];
                if (pnlStrict.some(p => concept === p || concept.includes(p))) return false;
            }
            
            if (isTypicalBalance || isNetIncomeInBalance) return true;

            return dataPeriods.some(p => {
                const curBIdx = p.balanceIdx !== -1 ? p.balanceIdx : p.pnlIdx;
                if (curBIdx === -1) return false;
                const val = getBalanceVal(row, curBIdx);
                return val !== 0 && !isNaN(val);
            });
        }).map(row => {
            const rawConcept = String(row[0]);
            let renamedConcept = rawConcept;
            const normConcept = normalizeText(rawConcept);
            
            const isTargetNetIncome = normConcept === "ganancia del periodo" || normConcept === "utilidad del ejercicio" || 
                normConcept === "resultado del periodo" || normConcept.includes("beneficio neto") || 
                normConcept.includes("utilidad neta") || normConcept.includes("ganancia neta") ||
                normConcept.includes("resultado neta");

            const rowValues = {};
            dataPeriods.forEach(p => {
                const curBIdx = p.balanceIdx !== -1 ? p.balanceIdx : p.pnlIdx;
                let val = curBIdx !== -1 ? getBalanceVal(row, curBIdx) : 0;
                
                if (isTargetNetIncome && val === 0 && curBIdx !== -1) {
                    const gAcum = getBalanceVal(balanceRows.gananciaAcumulada, curBIdx);
                    const uRet = getBalanceVal(balanceRows.utilidadesRetenidas, curBIdx);
                    if (gAcum !== 0 || uRet !== 0) val = uRet - gAcum;
                }
                
                rowValues[p.date.toLocaleDateString('es-ES', { month: 'short', year: 'numeric' })] = val;
            });

            if (isTargetNetIncome) renamedConcept = "Beneficio Neto del Periodo";
            return { concept: renamedConcept, values: rowValues };
        });

        const integrityGap = Math.abs(ingresos - costos - opex - ebitda);
        const integrityError = (ingresos !== 0) ? (integrityGap / Math.abs(ingresos)) > 0.01 : integrityGap > 100;

        const findRowVal = (rows, search) => {
            const r = rows.filter(r => r && r[0]).find(r => normalizeText(String(r[0])).includes(search));
            const curBIdx = bIdx !== -1 ? bIdx : pIdx;
            return (r && curBIdx !== -1) ? getBalanceVal(r, curBIdx) : 0;
        };

        const deudaTotal = findRowVal(bSheetToUse, "deuda total") || findRowVal(bSheetToUse, "deuda bruta");
        const ebitdaLTM = findRowVal(bSheetToUse, "ltm ebitda") || ebitda * 12;

        return {
            date: point.date.toLocaleDateString('es-ES', { month: 'short', year: 'numeric' }),
            sortDate: point.date,
            kpis: { 
                ingresos, 
                utilidad,
                ebitda, 
                margen_bruto: ingresos !== 0 ? (Math.abs(ingresos) - Math.abs(costos)) / Math.abs(ingresos) : 0,
                margen_ebitda: ingresos !== 0 ? ebitda / ingresos : 0, 
                margen_neto: ingresos !== 0 ? utilidad / ingresos : 0,
                cashflow: cashflowVal, // Este es el flujo neto
                cashEnding: cashflowDetail.ending || 0 // Este es el saldo final (Health)
            },
            balance: { 
                activos, pasivos, patrimonio, deudaTotal, ebitdaLTM,
                cuadra: Math.abs(activos - (pasivos + patrimonio)) < 100,
                fullRows: balanceFullRows 
            },
            cashflowDetail,
            integrity: { gap: integrityGap, isBroken: integrityError },
            tasaCambio: tasaCambio || 1,
            series: { ventas: [], ebitda: [] },
            pnl: { 
                categorias: { "Ingresos": ingresos, "Costo de Ventas": costos, "OPEX": opex, "EBITDA": ebitda, "Utilidad Neta": utilidad },
                opexDetalle: {
                    "Gastos Administrativos": pIdx !== -1 ? getVal(pnlRows.admin, pIdx) : 0,
                    "Gastos de Mercadeo": pIdx !== -1 ? getVal(pnlRows.mercadeo, pIdx) : 0,
                    "Gastos de Ventas (Comercial)": pIdx !== -1 ? getVal(pnlRows.comercial, pIdx) : 0,
                    "Gastos de Logística": pIdx !== -1 ? getVal(pnlRows.logistica, pIdx) : 0
                },
                segments,
                fullRows,
                detectedRows: {
                    ingresos: pnlRows.ingresos ? pnlRows.ingresos[0] : "No encontrada",
                    ebitda: pnlRows.ebitda ? pnlRows.ebitda[0] : "No encontrada",
                    costos: pnlRows.costos ? pnlRows.costos[0] : "No encontrada",
                    opex: pnlRows.opex ? pnlRows.opex[0] : "Calculado (Ventas - Costos - EBITDA)",
                    activos: balanceRows.activos ? balanceRows.activos[0] : "No detectado",
                    pasivo_patrimonio: (balanceRows.pasivos || balanceRows.patrimonio) ? `${balanceRows.pasivos?.[0] || ''} / ${balanceRows.patrimonio?.[0] || ''}` : "No detectado"
                }
            },
            alerts: ["FINANCIAL_STATEMENTS: Datos extraídos con desglose de segmentos."]
        };
    }).sort((a, b) => a.sortDate - b.sortDate);

    // Identificamos el cierre del año anterior (Diciembre 2025)
    const dic2025 = result.find(d => d.sortDate.getMonth() === 11 && d.sortDate.getFullYear() === 2025);
    
    // Tomamos los últimos 12 meses
    let finalSelection = result.slice(-12);
    
    // Si encontramos el cierre y NO está en la selección actual, lo inyectamos al principio
    if (dic2025 && !finalSelection.some(d => d.date === dic2025.date)) {
        finalSelection = [dic2025, ...finalSelection];
    }

    return { data: finalSelection };
}

function processWide(sheets) {
    const allRows = Object.values(sheets).flat();
    
    const getVal = (row, idx) => {
        if (!row || idx === undefined || idx === null) return 0;
        return cleanNumber(row[idx]);
    };

    const rowData = {
        ingresos: findRowByKeywords(allRows, FINANCIAL_KEYWORDS.ingresos),
        costos: findRowByKeywords(allRows, FINANCIAL_KEYWORDS.costos),
        opex: findRowByKeywords(allRows, FINANCIAL_KEYWORDS.opex),
        ebitda: findRowByKeywords(allRows, FINANCIAL_KEYWORDS.ebitda),
        utilidad: findRowByKeywords(allRows, FINANCIAL_KEYWORDS.utilidad),
        cashflow: findRowByKeywords(allRows, FINANCIAL_KEYWORDS.cashflow),
        tasa_cambio: findRowByKeywords(allRows, FINANCIAL_KEYWORDS.tasa_cambio)
    };

    const segmentKeywords = ["BT5", "EVP", "BON"];
    const segmentRows = detectSegments(allRows, segmentKeywords);

    if (!rowData.ingresos) return { error: "No se encontró una fila de 'Ingresos' o 'Ventas' en el reporte. Verifique que los nombres de las filas sean claros." };

    const sectionState = [];
    let lastSection = "monthly";
    for (let j = 0; j < (allRows[0]?.length || 0); j++) {
        let detected = null;
        for (let i = 0; i < Math.min(allRows.length, 5); i++) {
            const val = normalizeText(allRows[i]?.[j]);
            if (val.includes("ytd") || val.includes("acum") || val.includes("var")) { detected = "ytd"; break; }
            if (val.includes("monthly") || val.includes("mensual") || val.includes("real")) { detected = "monthly"; break; }
        }
        if (detected) lastSection = detected;
        sectionState[j] = lastSection;
    }

    let dataPoints = [];
    const monthNames = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"];
    for (let i = 0; i < Math.min(allRows.length, 30); i++) {
        const row = allRows[i];
        if (!row) continue;
        row.forEach((cell, j) => {
            let dateObj = null;
            if (cell instanceof Date) dateObj = cell;
            else if (typeof cell === 'number' && cell > 40000 && cell < 60000) dateObj = new Date((cell - 25569) * 86400 * 1000);
            else if (typeof cell === 'string') {
                const val = cell.toLowerCase();
                if (monthNames.some(m => val.includes(m))) {
                    dateObj = new Date();
                    const monthIdx = monthNames.findIndex(m => val.includes(m));
                    dateObj.setMonth(monthIdx);
                    if (val.match(/\d{4}/)) dateObj.setFullYear(parseInt(val.match(/\d{4}/)[0]));
                } else if (val.match(/^20\d{2}-\d{1,2}$/)) {
                    const [y, m] = val.split('-').map(Number);
                    dateObj = new Date(y, m - 1, 1);
                }
            }
            if (dateObj) {
                // Filtro de seguridad: No aceptar fechas más allá de 2026 (pedido por usuario)
                if (dateObj.getFullYear() >= 2022 && dateObj.getFullYear() <= 2026) {
                    if (sectionState[j] === "monthly") {
                        const val = cleanNumber(rowData.ingresos[j]);
                        if (val !== 0 && !dataPoints.some(p => p.idx === j)) {
                            dataPoints.push({ idx: j, date: dateObj });
                        }
                    }
                }
            }
        });
        if (dataPoints.length >= 2) break;
    }

    if (dataPoints.length === 0) {
        console.log("🔍 No se detectaron fechas en Wide Format, intentando fallback...");
        rowData.ingresos.forEach((cell, j) => {
            if (j === 0) return;
            const val = cleanNumber(cell);
            if (val !== 0 && !isNaN(val)) {
                if (!dataPoints.some(p => p.idx === j)) {
                    const d = new Date();
                    d.setMonth(d.getMonth() - (rowData.ingresos.length - j));
                    dataPoints.push({ idx: j, date: d });
                }
            }
        });
    }

    if (dataPoints.length === 0) {
        return { error: "No se encontraron datos numéricos en las columnas del reporte gerencial." };
    }

    const result = dataPoints.map(point => {
        const ingresos = cleanNumber(rowData.ingresos[point.idx]);
        const costos = rowData.costos ? cleanNumber(rowData.costos[point.idx]) : 0;
        const opex = rowData.opex ? cleanNumber(rowData.opex[point.idx]) : 0;
        const ebitda = rowData.ebitda ? cleanNumber(rowData.ebitda[point.idx]) : (ingresos - costos - opex);
        const utilidad = rowData.utilidad ? cleanNumber(rowData.utilidad[point.idx]) : 0;
        const cashflow = rowData.cashflow ? cleanNumber(rowData.cashflow[point.idx]) : utilidad;
        const tasaCambio = rowData.tasa_cambio ? cleanNumber(rowData.tasa_cambio[point.idx]) : 1;

        const segments = {};
        Object.entries(segmentRows).forEach(([name, data]) => {
            const sumVals = (rowList) => rowList.reduce((acc, row) => acc + cleanNumber(row[point.idx]), 0);
            segments[name] = {
                ventas: sumVals(data.ventasRows),
                costos: sumVals(data.costosRows)
            };
        });

        // Capturar todas las filas del P&L para la vista detallada
        const fullRows = allRows.filter(row => {
            if (!row) return false;
            let conceptRaw = row[0];
            if (!conceptRaw || String(conceptRaw).trim().toUpperCase() === 'X') conceptRaw = row[1];
            if (!conceptRaw) return false;
            const concept = normalizeText(String(conceptRaw));
            if (concept.includes("formatcode") || concept.includes("unnamed") || concept.length < 2) return false;

            // Filtramos filas que tengan al menos 1 número en los dataPoints, o si son Categorias (sin numeros)
            // Agregamos también las filas que sean categorias (por ejemplo "Estado de Resultados") aunque no tengan numeros
            const isCategory = (concept === "estado de resultados" || concept === "estado de situacion" || concept === "kpis y drivers" || concept === "modulo deuda" || concept === "analisis horizontal" || concept === "analisis vertical" || concept === "analisis margen" || concept === "rentabilidad" || concept === "variables macro" || concept === "balances deuda" || concept === "schedule amortizacion" || concept === "kpis deuda");
            return isCategory || dataPoints.some(p => typeof row[p.idx] === 'number' || (!isNaN(cleanNumber(row[p.idx])) && cleanNumber(row[p.idx]) !== 0));
        }).map(row => {
            const rowValues = {};
            dataPoints.forEach(p => {
                rowValues[p.date.toLocaleDateString('es-ES', { month: 'short', year: 'numeric' })] = cleanNumber(row[p.idx]);
            });
            let conceptRaw = row[0];
            if (!conceptRaw || String(conceptRaw).trim().toUpperCase() === 'X') conceptRaw = row[1];
            const rawConcept = String(conceptRaw).trim();
            const renamedConcept = (normalizeText(rawConcept) === "ventas p6") ? "Ventas BON" : rawConcept;

            return {
                concept: renamedConcept,
                values: rowValues
            };
        });

        return {
            date: point.date.toLocaleDateString('es-ES', { month: 'short', year: 'numeric' }),
            sortDate: point.date,
            kpis: {
                ingresos,
                utilidad,
                ebitda,
                margen_bruto: ingresos !== 0 ? (Math.abs(ingresos) - Math.abs(costos)) / Math.abs(ingresos) : 0,
                margen_ebitda: ingresos !== 0 ? ebitda / ingresos : 0,
                margen_neto: ingresos !== 0 ? utilidad / ingresos : 0,
                cashflow
            },
            balance: { activos: 0, pasivos: 0, patrimonio: 0, cuadra: true },
            tasaCambio: tasaCambio,
            series: { ventas: [], ebitda: [] },
            pnl: { 
                categorias: { "Ingresos": ingresos, "Costo de Ventas": costos, "OPEX": opex, "EBITDA": ebitda, "Utilidad Neta": utilidad },
                segments: segments,
                fullRows: fullRows,
                detectedRows: {
                    ingresos: rowData.ingresos ? rowData.ingresos[0] : "No encontrada",
                    ebitda: rowData.ebitda ? rowData.ebitda[0] : "No encontrada",
                    costos: rowData.costos ? rowData.costos[0] : "No encontrada"
                }
            },
            alerts: ["WIDE_FORMAT: Reporte gerencial detectado."]
        };
    }).sort((a, b) => a.sortDate - b.sortDate);

    // Final deduplication by date string
    const uniqueResult = [];
    const seenDates = new Set();
    result.forEach(item => {
        if (!seenDates.has(item.date)) {
            seenDates.add(item.date);
            uniqueResult.push(item);
        }
    });

    return { data: uniqueResult.slice(-12) };
}

function processTBSetup(sheets, tbKey, setupKey) {
    const tbSheet = sheets[tbKey];
    const setupSheet = sheets[setupKey];
    
    const setupMap = new Map();
    const diagnostics = { 
        rows: 0, 
        mapped: 0, 
        tbSample: [], 
        setupSample: [],
        tbColDetected: -1,
        setupColDetected: -1
    };

    function cleanAccount(val) {
        if (val === undefined || val === null) return "";
        return String(val).trim().toLowerCase().replace(/[^a-z0-9]/g, '');
    }

    function detectAccountColumnByContent(sheet, startRow) {
        const scores = [];
        const sampleRows = sheet.slice(startRow, startRow + 150);
        sampleRows.forEach(row => {
            if (!row) return;
            row.forEach((cell, j) => {
                if (scores[j] === undefined) scores[j] = 0;
                const val = String(cell).trim();
                if (!val || val.length < 2) return;
                if (/^\d{4,10}$/.test(val)) scores[j] += 10;
                else if (/^(\d+[\.\-])+\d+$/.test(val)) scores[j] += 15;
                else if (/^[A-Z0-9]{4,12}$/i.test(val) && !val.includes(" ")) scores[j] += 5;
            });
        });
        let bestCol = -1; let maxScore = 0;
        scores.forEach((score, j) => { if (score > maxScore) { maxScore = score; bestCol = j; } });
        return { col: bestCol, score: maxScore };
    }

    function detectDateColumnByContent(sheet, startRow) {
        const scores = [];
        const sampleRows = sheet.slice(startRow, startRow + 100);
        const monthNames = ["ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic", "jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"];
        sampleRows.forEach(row => {
            if (!row) return;
            row.forEach((cell, j) => {
                if (scores[j] === undefined) scores[j] = 0;
                if (cell instanceof Date) scores[j] += 20;
                else if (typeof cell === 'number' && cell > 35000 && cell < 60000) scores[j] += 15;
                else if (typeof cell === 'string') {
                    const val = cell.toLowerCase();
                    if (monthNames.some(m => val.includes(m))) scores[j] += 10;
                    if (/\d{4}[-\/]\d{2}/.test(val)) scores[j] += 12;
                }
            });
        });
        let bestCol = -1; let maxScore = 0;
        scores.forEach((score, j) => { if (score > maxScore) { maxScore = score; bestCol = j; } });
        return { col: bestCol, score: maxScore };
    }

    function detectBalanceColumnByContent(sheet, startRow) {
        const scores = [];
        const sampleRows = sheet.slice(startRow, startRow + 150);
        sampleRows.forEach(row => {
            if (!row) return;
            row.forEach((cell, j) => {
                if (scores[j] === undefined) scores[j] = 0;
                const val = Number(cell);
                if (isNaN(val) || cell === null || cell === "") return;
                if (!Number.isInteger(val)) scores[j] += 10;
                if (Math.abs(val) > 100) scores[j] += 5;
            });
        });
        let bestCol = -1; let maxScore = 0;
        scores.forEach((score, j) => { if (score > maxScore) { maxScore = score; bestCol = j; } });
        return { col: bestCol, score: maxScore };
    }

    function findHeaderRow(sheet, keywords) {
        let bestRow = 0; let maxScore = -1; let bestCols = {};
        for (let i = 0; i < Math.min(sheet.length, 30); i++) {
            const row = sheet[i]; if (!row) continue;
            let currentScore = 0; let currentCols = {};
            row.forEach((cell, j) => {
                const c = normalizeText(cell); if (!c) return;
                for (const [key, searchTerms] of Object.entries(keywords)) {
                    if (searchTerms.some(term => c === term || (term.length > 3 && c.includes(term)))) {
                        if (!currentCols[key]) { currentCols[key] = j; currentScore++; }
                    }
                }
            });
            if (currentScore > maxScore) { maxScore = currentScore; bestRow = i; bestCols = currentCols; }
        }
        return { row: bestRow, cols: bestCols, score: maxScore };
    }

    const setupKeywords = {
        cuenta: ["cuenta", "codigo", "acct", "account", "cta", "id"],
        cat: ["categoria", "grupo", "category", "clase", "tipo"],
        sub: ["subcategoria", "subcat", "subgrupo"],
        signo: ["signo", "multiplicador", "sign", "naturaleza", "factor"]
    };

    const setupHeader = findHeaderRow(setupSheet, setupKeywords);
    let setupCols = setupHeader.cols;
    let setupHeaderRow = setupHeader.row;

    const setupContentAcc = detectAccountColumnByContent(setupSheet, setupHeaderRow + 1);
    if (setupCols.cuenta === undefined || setupCols.cuenta === -1) setupCols.cuenta = setupContentAcc.col !== -1 ? setupContentAcc.col : 0;
    if (setupCols.cat === undefined || setupCols.cat === -1) setupCols.cat = 1;
    if (setupCols.sub === undefined || setupCols.sub === -1) setupCols.sub = 2;
    if (setupCols.signo === undefined || setupCols.signo === -1) setupCols.signo = setupSheet[setupHeaderRow + 1] ? setupSheet[setupHeaderRow + 1].length - 1 : 3;

    diagnostics.setupColDetected = setupCols.cuenta;

    setupSheet.forEach((row, i) => {
        if (i <= setupHeaderRow || !row) return;
        const cuenta = cleanAccount(row[setupCols.cuenta]);
        if (!cuenta) return;
        if (diagnostics.setupSample.length < 5) diagnostics.setupSample.push(cuenta);
        setupMap.set(cuenta, {
            categoria: row[setupCols.cat] || "Sin Categoría",
            subcategoria: row[setupCols.sub] || "Sin Subcategoría",
            signo: cleanNumber(row[setupCols.signo]) || 1
        });
    });

    if (setupMap.size === 0) return { error: "La hoja 'Setup' no tiene datos de cuenta válidos." };

    const tbKeywords = {
        cuenta: ["cuenta", "codigo", "acct", "account", "cta", "id", "cod"],
        fecha: ["fecha", "periodo", "mes", "date", "year", "ano", "time", "fec"],
        balance: ["balance", "saldo", "monto", "final", "amount", "debe", "haber", "neto", "total", "valor"]
    };

    const tbHeader = findHeaderRow(tbSheet, tbKeywords);
    let tbCols = tbHeader.cols;
    let tbHeaderRow = tbHeader.row;

    const tbContentAcc = detectAccountColumnByContent(tbSheet, tbHeaderRow + 1);
    const tbContentDate = detectDateColumnByContent(tbSheet, tbHeaderRow + 1);
    const tbContentBal = detectBalanceColumnByContent(tbSheet, tbHeaderRow + 1);

    if (tbCols.cuenta === undefined || tbCols.cuenta === -1) tbCols.cuenta = tbContentAcc.col !== -1 ? tbContentAcc.col : 0;
    if (tbCols.fecha === undefined || tbCols.fecha === -1) tbCols.fecha = tbContentDate.col !== -1 ? tbContentDate.col : 1;
    if (tbCols.balance === undefined || tbCols.balance === -1) tbCols.balance = tbContentBal.col !== -1 ? tbContentBal.col : 2;

    diagnostics.tbColDetected = tbCols.cuenta;

    const monthlyAggregates = {};

    tbSheet.forEach((row, idx) => {
        if (idx <= tbHeaderRow || !row) return;
        const cuenta = cleanAccount(row[tbCols.cuenta]);
        if (!cuenta) return;
        diagnostics.rows++;
        if (diagnostics.tbSample.length < 5 && !diagnostics.tbSample.includes(cuenta)) diagnostics.tbSample.push(cuenta);

        const setup = setupMap.get(cuenta);
        if (!setup) return;
        diagnostics.mapped++;

        const rawDate = row[tbCols.fecha];
        let dateObj = null;
        if (rawDate instanceof Date) dateObj = rawDate;
        else if (typeof rawDate === 'number') dateObj = new Date((rawDate - 25569) * 86400 * 1000);
        else if (typeof rawDate === 'string') dateObj = new Date(rawDate);

        if (!dateObj || isNaN(dateObj.getTime())) return;

        const dateKey = dateObj.toLocaleDateString('es-ES', { month: 'short', year: 'numeric' });
        if (!monthlyAggregates[dateKey]) {
            monthlyAggregates[dateKey] = { 
                kpis: { ingresos: 0, utilidad: 0, ebitda: 0, margen_bruto: 0, margen_ebitda: 0, margen_neto: 0, cashflow: 0 },
                balance: { activos: 0, pasivos: 0, patrimonio: 0, cuadra: false },
                series: { ventas: [], ebitda: [] },
                pnl: { categorias: {} },
                alerts: [],
                sortDate: dateObj,
                date: dateKey,
                _raw: { ingresos: 0, costos: 0, gastos: 0 }
            };
        }

        const valorAjustado = cleanNumber(row[tbCols.balance]) * setup.signo;
        const cat = normalizeText(setup.categoria);
        const agg = monthlyAggregates[dateKey];

        if (cat.includes("ingreso") || cat.includes("venta")) agg._raw.ingresos += valorAjustado;
        else if (cat.includes("costo")) agg._raw.costos += valorAjustado;
        else if (cat.includes("gasto")) agg._raw.gastos += valorAjustado;

        if (cat.includes("activo")) agg.balance.activos += valorAjustado;
        else if (cat.includes("pasivo")) agg.balance.pasivos += valorAjustado;
        else if (cat.includes("patrimonio")) agg.balance.patrimonio += valorAjustado;

        if (!agg.pnl.categorias[setup.categoria]) agg.pnl.categorias[setup.categoria] = 0;
        agg.pnl.categorias[setup.categoria] += valorAjustado;
    });

    const result = Object.values(monthlyAggregates).sort((a, b) => a.sortDate - b.sortDate).map(agg => {
        agg.kpis.ingresos = agg._raw.ingresos;
        agg.kpis.ebitda = agg._raw.ingresos - agg._raw.costos - agg._raw.gastos;
        agg.kpis.margen_bruto = agg.kpis.ingresos !== 0 ? (Math.abs(agg.kpis.ingresos) - Math.abs(agg._raw.costos)) / Math.abs(agg.kpis.ingresos) : 0;
        agg.kpis.margen_ebitda = agg.kpis.ingresos !== 0 ? agg.kpis.ebitda / agg.kpis.ingresos : 0;
        agg.kpis.cashflow = agg.kpis.ebitda;
        agg.balance.cuadra = Math.abs(agg.balance.activos - (agg.balance.pasivos + agg.balance.patrimonio)) < 100;
        return agg;
    });

    if (result.length === 0) {
        return { error: `ERROR DE MAPEADO: No hay coincidencia entre TB y Setup.\n\nDIAGNÓSTICO:\n${JSON.stringify(diagnostics, null, 2)}` };
    }

    return { data: result };
}
