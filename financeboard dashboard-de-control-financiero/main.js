import * as XLSX from 'xlsx';
import { GoogleGenAI } from "@google/genai";
import * as d3 from 'd3';
import { financialEngine, formatCurrency, formatRawCurrency, formatPercent, normalizeText } from "./financialEngine.js";
import { buildLLMInput } from "./buildLLMInput.js";
import { validateLLMInput } from "./validator.js";

let globalFinancialData = [];
const loader = document.getElementById('loader');
const monthSelector = document.getElementById('monthSelector');

// Initialize Gemini
const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

// MSAL Configuration
const msalConfig = {
    auth: {
        clientId: import.meta.env.VITE_MICROSOFT_CLIENT_ID || "cd40e757-85f4-4676-89ec-78445851aa92",
        authority: "https://login.microsoftonline.com/8dbe3e04-118c-4cd5-ae67-0c0c21606098",
        redirectUri: window.location.origin,
    },
    cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false,
    }
};

let msalInstance;
if (window.msal) {
    msalInstance = new window.msal.PublicClientApplication(msalConfig);
}

const SHARPOINT_FILE_URL = "https://aguaplanetaazul2-my.sharepoint.com/personal/marcos_ojeda_planetaazulrd_com/_layouts/15/Doc.aspx?sourcedoc={cfe13828-c964-447a-8147-feb8de79816c}&download=1";

async function connectM365() {
    if (!msalInstance) {
        alert("MSAL no inicializado.");
        return;
    }

    try {
        // En versiones recientes, debemos asegurarnos del estado local si usamos msal
        await msalInstance.initialize?.(); 
        await msalInstance.handleRedirectPromise?.();

        const loginResponse = await msalInstance.loginPopup({
            scopes: ["User.Read", "Files.Read", "Files.Read.All"],
            prompt: "select_account"
        });
        const token = loginResponse.accessToken;
        
        await fetchMasterData(token);
    } catch (error) {
        if (error.errorCode === "interaction_in_progress") {
            // Intenta limpiar el estado en caso de que se haya quedado pegado
            alert("Hay una autenticación en progreso o el popup fue bloqueado. Por favor, recargue la página (o ábrala en una nueva pestaña) e intente de nuevo.");
            return;
        }
        if (error.message && error.message.includes("popup_window_error")) {
             alert("El navegador bloqueó la ventana emergente. Por favor, asegúrese de abrir esta aplicación en una NUEVA PESTAÑA completa, o permita los popups para este sitio.");
             return;
        }
        console.error(error);
        alert("Error autenticando con Office 365: " + error.message);
    }
}

async function fetchMasterData(token = null) {
    const statusEl = document.getElementById('engineStatus');
    const sidebarSyncDot = document.getElementById('sidebarSyncDot');
    const sidebarSyncText = document.getElementById('sidebarSyncText');

    if (sidebarSyncDot) sidebarSyncDot.style.backgroundColor = 'var(--warning)';
    if (sidebarSyncText) {
        sidebarSyncText.innerText = 'Sincronizando...';
        sidebarSyncText.style.color = 'var(--warning)';
    }

    if (statusEl) {
        statusEl.style.background = '#e0f2fe';
        statusEl.style.color = '#0369a1';
        statusEl.style.borderColor = '#bae6fd';
        statusEl.innerHTML = "⏳ Sincronizando modelo remoto...";
    }
    if (loader) loader.style.display = 'flex';
    
    const loginBtn = document.getElementById('loginM365Btn');
    if (loginBtn) loginBtn.style.display = 'none';

    try {
        let arrayBuffer;
        
        if (token) {
            // Attempt generic Graph API request by encoding the sharing URL
            const encodedUrl = btoa(SHARPOINT_FILE_URL).replace(/=/g, '').replace(/\//g, '_').replace(/\+/g, '-');
            const graphUrl = `https://graph.microsoft.com/v1.0/shares/u!${encodedUrl}/driveItem/content`;
            
            const req = await fetch(graphUrl, {
                headers: { "Authorization": `Bearer ${token}` }
            });
            if (!req.ok) throw new Error(`O365 Graph Error: ${req.status} ${req.statusText}`);
            arrayBuffer = await req.arrayBuffer();
        } else {
            // Unauthenticated fallback proxy approach
            const response = await fetch("/api/downloadSync");
            if (!response.ok) {
                const errData = await response.json().catch(() => ({}));
                throw new Error(errData.error || `Proxy Error: ${response.status}`);
            }
            arrayBuffer = await response.arrayBuffer();
            
            // Check if what we got is an HTML login page instead of an Excel file
            const uint8Array = new Uint8Array(arrayBuffer);
            // Check for '<html' or '<!DOC' at the beginning
            const textHead = new TextDecoder().decode(uint8Array.slice(0, 100)).toLowerCase();
            if (textHead.includes('<html') || textHead.includes('<!doc')) {
                 throw new Error("El enlace es privado y redirigió a la página de inicio de sesión de Microsoft. Debe iniciar sesión con Office 365 o cargar el archivo manualmente.");
            }
        }
        
        const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array', cellDates: true });
        const engineResult = financialEngine(workbook);
        
        if (engineResult.error || !engineResult.data || engineResult.data.length === 0) {
            throw new Error(engineResult.error || "No se pudieron extraer datos numéricos del archivo.");
        }
        
        if (statusEl) {
            statusEl.innerHTML = "✅ Sincronizado con O365";
        }
        
        if (sidebarSyncDot) sidebarSyncDot.style.backgroundColor = 'var(--success)';
        if (sidebarSyncText) {
            sidebarSyncText.innerText = 'Sincronizado';
            sidebarSyncText.style.color = 'var(--success)';
        }
        
        globalFinancialData = engineResult.data;
        renderDashboard(globalFinancialData);
        if (loader) loader.style.display = 'none';
        
    } catch (error) {
        if (error.message && error.message.includes("El enlace es privado")) {
            console.warn("Auto-sync fallback triggered (expected):", error.message);
        } else {
            console.error("Auto-sync failed:", error);
        }
        if (statusEl) {
            statusEl.style.background = '#fee2e2';
            statusEl.style.color = '#991b1b';
            statusEl.style.borderColor = '#fecaca';
            statusEl.innerHTML = "⚠️ Sincronización fallida. Presione 'Conectar Office 365' o use la carga manual.";
            statusEl.title = error.message; 
        }
        
        if (sidebarSyncDot) sidebarSyncDot.style.backgroundColor = 'var(--danger)';
        if (sidebarSyncText) {
             sidebarSyncText.innerText = 'Desconectado';
             sidebarSyncText.style.color = 'var(--danger)';
        }
        
        if (loader) loader.style.display = 'none';
        
        if (loginBtn) loginBtn.style.display = 'flex'; // Show login button
    }
}

document.addEventListener('DOMContentLoaded', () => {
    // If MSAL accounts exist, we could silently login, but for now we require button
    fetchMasterData();
    
    const loginM365Btn = document.getElementById('loginM365Btn');
    if (loginM365Btn) {
        loginM365Btn.addEventListener('click', connectM365);
    }
    const fileInput = document.getElementById('fileInput');
    const dropZone = document.getElementById('dropZone');

    // Setup Export and Mobile Menu
    const btnExportExcel = document.getElementById('btnExportExcel');
    if (btnExportExcel) {
        btnExportExcel.addEventListener('click', () => {
            if (!globalFinancialData || globalFinancialData.length === 0) {
                alert('No hay datos para exportar.');
                return;
            }
            exportToExcelSuite(globalFinancialData);
        });
    }

    const btnExportPDF = document.getElementById('btnExportPDF');
    if (btnExportPDF) {
        btnExportPDF.addEventListener('click', () => {
            if (!globalFinancialData || globalFinancialData.length === 0) {
                alert('No hay datos para exportar.');
                return;
            }
            
            const mainContainer = document.querySelector('.main-container');
            const views = mainContainer.querySelectorAll('.view-container');
            const headerActions = document.querySelector('.header-actions');
            const sidebar = document.querySelector('.sidebar');
            const mobileHeader = document.querySelector('.mobile-header');
            
            // Store original state
            let activeViewId = '';
            views.forEach(v => {
                if (v.classList.contains('active')) activeViewId = v.id;
            });
            const originalHeaderDisplay = headerActions ? headerActions.style.display : '';
            const originalSidebarDisplay = sidebar ? sidebar.style.display : '';
            const originalMobileHeaderDisplay = mobileHeader ? mobileHeader.style.display : '';
            const originalMainPadding = mainContainer.style.padding;
            const originalOverflow = mainContainer.style.overflow;
            
            // Force charts visibility for PDF
            const dashboardChartsGrid = document.querySelector('.dashboard-charts-grid');
            const originalChartsGridDisplay = dashboardChartsGrid ? dashboardChartsGrid.style.display : '';
            if (dashboardChartsGrid) {
                dashboardChartsGrid.style.setProperty('display', 'grid', 'important');
            }

            // Modify DOM for PDF capture
            if (headerActions) headerActions.style.display = 'none';
            if (sidebar) sidebar.style.display = 'none';
            if (mobileHeader) mobileHeader.style.display = 'none';
            
            mainContainer.style.padding = '20px';
            mainContainer.style.overflow = 'visible';

            views.forEach(v => {
                if (v.id !== 'view-config') {
                    v.classList.add('active');
                    v.style.display = 'block';
                    v.style.pageBreakAfter = 'always';
                } else {
                    v.style.display = 'none';
                }
            });

            // Trigger resize event to force D3 to redraw if necessary.
            window.dispatchEvent(new Event('resize'));
            
            const opt = {
                margin:       [0.5, 0.5, 0.5, 0.5],
                filename:     'Reporte_Planeta_Azul.pdf',
                image:        { type: 'jpeg', quality: 0.98 },
                html2canvas:  { scale: 2, useCORS: true, logging: false, windowWidth: 1200 },
                jsPDF:        { unit: 'in', format: 'letter', orientation: 'portrait' }
            };

            // Wait 800ms before generating PDF
            setTimeout(() => {
                html2pdf().set(opt).from(mainContainer).save().then(() => {
                    // Restore original state
                    if (headerActions) headerActions.style.display = originalHeaderDisplay;
                    if (sidebar) sidebar.style.display = originalSidebarDisplay;
                    if (mobileHeader) mobileHeader.style.display = originalMobileHeaderDisplay;
                    if (dashboardChartsGrid) dashboardChartsGrid.style.display = originalChartsGridDisplay;
                    
                    mainContainer.style.padding = originalMainPadding;
                    mainContainer.style.overflow = originalOverflow;
                    
                    views.forEach(v => {
                        v.style.display = '';
                        v.style.pageBreakAfter = '';
                        if (v.id !== activeViewId) {
                            v.classList.remove('active');
                        } else {
                            v.classList.add('active');
                        }
                    });
                    
                    window.dispatchEvent(new Event('resize'));
                });
            }, 800);
        });
    }

    const menuToggleBtn = document.getElementById('menuToggleBtn');
    const sidebar = document.querySelector('.sidebar');
    if (menuToggleBtn && sidebar) {
        menuToggleBtn.addEventListener('click', () => {
            sidebar.classList.toggle('open');
        });
    }

    if (fileInput) {
        fileInput.addEventListener('change', handleFileUpload);
    }

    if (dropZone) {
        dropZone.addEventListener('dragover', (e) => {
            e.preventDefault();
            dropZone.style.borderColor = 'var(--primary)';
            dropZone.style.background = 'rgba(37, 99, 235, 0.05)';
        });

        dropZone.addEventListener('dragleave', () => {
            dropZone.style.borderColor = 'rgba(255,255,255,0.2)';
            dropZone.style.background = 'rgba(255,255,255,0.05)';
        });

        dropZone.addEventListener('drop', (e) => {
            e.preventDefault();
            dropZone.style.borderColor = 'rgba(255,255,255,0.2)';
            dropZone.style.background = 'rgba(255,255,255,0.05)';
            const files = e.dataTransfer.files;
            if (files.length > 0) {
                fileInput.files = files;
                handleFileUpload({ target: { files } });
            }
        });
    }

    if (monthSelector) {
        monthSelector.addEventListener('change', (e) => {
            const index = parseInt(e.target.value);
            if (!isNaN(index)) updateUI(globalFinancialData, index);
        });
    }

    // Navigation Logic
    const menuItems = ['menu-kpi', 'menu-resumen', 'menu-pnl', 'menu-balance', 'menu-cashflow', 'menu-estados', 'menu-config', 'menu-glosario'];
    menuItems.forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.addEventListener('click', (e) => {
                e.preventDefault();
                // Update active class
                menuItems.forEach(mId => document.getElementById(mId)?.classList.remove('active'));
                el.classList.add('active');

                // Switch views
                document.querySelectorAll('.view-container').forEach(v => v.classList.remove('active'));
                const viewId = id.replace('menu-', 'view-');
                document.getElementById(viewId)?.classList.add('active');

                // Close mobile sidebar if open
                const sidebar = document.querySelector('.sidebar');
                if (sidebar && window.innerWidth <= 1024) {
                    sidebar.classList.remove('open');
                }

                // Update Title
                const titleLabel = document.getElementById('titleLabel');
                const titles = {
                    'menu-kpi': "Torre de Control: Indicadores Clave",
                    'menu-resumen': "Dashboard de Gestión Corporativa (RD$)",
                    'menu-pnl': "Estado de Resultados Detallado (RD$)",
                    'menu-balance': "Balance General Consolidado (RD$)",
                    'menu-cashflow': "Estado de Flujo de Efectivo (RD$)",
                    'menu-config': "Configuración y Auditoría",
                    'menu-estados': "Estados Financieros",
                    'menu-glosario': "Glosario de Términos y Metodologías Financieras"
                };
                if (titles[id]) titleLabel.textContent = titles[id];

                const periodContainer = document.getElementById('periodContainer');
                if (periodContainer) {
                    if (id === 'menu-glosario' || id === 'menu-config') {
                        periodContainer.style.display = 'none';
                    } else {
                        periodContainer.style.display = 'flex';
                    }
                }

                const searchWrapper = document.getElementById('searchContainerWrapper');
                if (monthSelector) {
                    if (id === 'menu-config' || id === 'menu-glosario') {
                        monthSelector.style.display = 'none';
                    } else if (globalFinancialData && globalFinancialData.length > 0) {
                        monthSelector.style.display = 'block';
                    }
                }
                
                if (searchWrapper) {
                    const viewsWithSearch = ['menu-resumen', 'menu-pnl', 'menu-balance', 'menu-cashflow', 'menu-estados'];
                    if (viewsWithSearch.includes(id) && globalFinancialData && globalFinancialData.length > 0) {
                        searchWrapper.style.display = 'flex';
                    } else {
                        searchWrapper.style.display = 'none';
                    }
                }

                // Re-render UI so that D3 charts and things calculate dimensions properly
                // when this view becomes visible.
                if (globalFinancialData && globalFinancialData.length > 0 && monthSelector) {
                    const idx = parseInt(monthSelector.value);
                    if (!isNaN(idx)) updateUI(globalFinancialData, idx);
                }
            });
        }
    });

    const accountSearch = document.getElementById('accountSearch');
    if (accountSearch) {
        accountSearch.addEventListener('input', (e) => {
            const query = e.target.value.toLowerCase();
            
            // Filter desktop tables
            const tablesToFilter = ['pnlDetailedTable', 'balanceTable', 'covenantTable', 'cashflowTable', 'cfMetricsTable', 'table-estados', 'tableResumenOperativo', 'tableVentasSegmento', 'tableCostosSegmento', 'tableOpex'];
            tablesToFilter.forEach(tId => {
                const table = document.getElementById(tId);
                if (table) {
                    const rows = table.querySelectorAll('tbody tr');
                    rows.forEach(tr => {
                        const firstCell = tr.querySelector('td:first-child');
                        if (firstCell) {
                            const accountName = firstCell.textContent.toLowerCase();
                            if (accountName.includes(query)) {
                                tr.style.display = '';
                            } else {
                                tr.style.display = 'none';
                            }
                        }
                    });
                }
            });

            // Filter mobile cards
            const mobileContainersToFilter = [
                 'pnlMobileContainer', 'balanceMobileContainer', 'covenantMobileContainer', 
                 'cashflowMobileContainer', 'cfMetricsMobileContainer', 'estadosMobileContainer',
                 'resumenOperativoMobileContainer', 'ventasSegmentoMobileContainer', 'costosSegmentoMobileContainer', 'opexMobileContainer'
            ];
            mobileContainersToFilter.forEach(cId => {
                const container = document.getElementById(cId);
                if (container) {
                    const cards = container.querySelectorAll('.mobile-vertical-card');
                    cards.forEach(card => {
                        const titleEl = card.querySelector('.mobile-vertical-card-title span');
                        if (titleEl) {
                            const accountName = titleEl.textContent.toLowerCase();
                            if (accountName.includes(query)) {
                                card.style.display = '';
                            } else {
                                card.style.display = 'none';
                            }
                        }
                    });

                    // Hide empty accordion groups
                    const accordions = container.querySelectorAll('.mobile-accordion-group');
                    accordions.forEach(acc => {
                        const visibleCards = acc.querySelectorAll('.mobile-vertical-card[style=""]');
                        // if searching and no visible cards, hide the whole accordion
                        if (query !== '' && visibleCards.length === 0 && acc.querySelectorAll('.mobile-vertical-card').length > 0) {
                            acc.style.display = 'none';
                        } else {
                            acc.style.display = '';
                            if (query !== '') {
                                // Auto expand if searching
                                const content = acc.querySelector('.mobile-accordion-content');
                                if (content) content.classList.add('open');
                            }
                        }
                    });
                }
            });
        });
    }

    if (typeof lucide !== 'undefined') lucide.createIcons();
    
    // Global polished tooltip system for KPI Cards (matches chart style)
    let globalTooltip = d3.select("body").select(".d3-tooltip");
    if (globalTooltip.empty()) {
        globalTooltip = d3.select("body")
            .append("div")
            .attr("class", "d3-tooltip")
            .style("opacity", 0);
    }

    // Add event delegation for any element with data-tooltip
    document.addEventListener('mouseover', (e) => {
        const trigger = e.target.closest('[data-tooltip]');
        if (trigger) {
            const text = trigger.getAttribute('data-tooltip');
            globalTooltip.style("opacity", 1)
                .html(text);
        }
    });

    document.addEventListener('mousemove', (e) => {
        const trigger = e.target.closest('[data-tooltip]');
        if (trigger) {
            globalTooltip
                .style("left", (e.pageX + 15) + "px")
                .style("top", (e.pageY - 15) + "px");
        }
    });

    document.addEventListener('mouseout', (e) => {
        const trigger = e.target.closest('[data-tooltip]');
        if (trigger) {
            globalTooltip.style("opacity", 0);
        }
    });

    // Support for touch devices (click to show/hide)
    document.addEventListener('click', (e) => {
        const trigger = e.target.closest('[data-tooltip]');
        if (trigger && window.innerWidth < 1024) {
            const isVisible = globalTooltip.style("opacity") === "1";
            if (isVisible) {
                globalTooltip.style("opacity", 0);
            } else {
                const text = trigger.getAttribute('data-tooltip');
                globalTooltip.style("opacity", 1)
                    .html(text)
                    .style("left", (e.pageX + 15) + "px")
                    .style("top", (e.pageY - 15) + "px");
            }
        } else if (!trigger) {
            globalTooltip.style("opacity", 0);
        }
    });

    // Go to top button logic
    const mainContainer = document.querySelector('.main-container');
    const scrollTopBtn = document.getElementById('scrollTopBtn');
    if (mainContainer && scrollTopBtn) {
        mainContainer.addEventListener('scroll', () => {
            if (mainContainer.scrollTop > 300) {
                scrollTopBtn.classList.add('visible');
            } else {
                scrollTopBtn.classList.remove('visible');
            }
        });
        
        // Mobile fallback for body scroll
        window.addEventListener('scroll', () => {
             if (window.scrollY > 300) {
                 scrollTopBtn.classList.add('visible');
             } else {
                 scrollTopBtn.classList.remove('visible');
             }
        });

        scrollTopBtn.addEventListener('click', () => {
            mainContainer.scrollTo({ top: 0, behavior: 'smooth' });
            window.scrollTo({ top: 0, behavior: 'smooth' });
        });
    }

    // Handle window resize for D3 Charts redrawing and Mobile Accordions
    window.addEventListener('resize', () => {
        if (globalFinancialData && globalFinancialData.length > 0 && monthSelector) {
            const idx = parseInt(monthSelector.value);
            if (!isNaN(idx)) {
                // Throttle maybe not strictly needed for this scale, but good practice
                const rollingData = globalFinancialData.slice(Math.max(0, idx - 11), idx + 1).filter(d => !isYear2025(d));
                renderMarginChart(rollingData);
                renderCashFlowChart(rollingData);
                
                // Rebuild Mobile Accordions if crossing breakpoint
                buildMobileAccordionsFromTable('pnlDetailedTable', 'pnlMobileContainer');
                buildMobileAccordionsFromTable('balanceTable', 'balanceMobileContainer');
                buildMobileAccordionsFromTable('covenantTable', 'covenantMobileContainer');
                buildMobileAccordionsFromTable('cashflowTable', 'cashflowMobileContainer');
                buildMobileAccordionsFromTable('cfMetricsTable', 'cfMetricsMobileContainer');
                buildMobileAccordionsFromTable('table-estados', 'estadosMobileContainer');
                
                // Resumen
                const lastData = globalFinancialData[idx];
                const kpis = lastData.kpis || { ingresos: 0, ebitda: 0, margen_ebitda: 0 };
                const categories = (lastData.pnl && lastData.pnl.categorias) ? lastData.pnl.categorias : {};
                const totalCost = categories["Costo de Ventas"] || 0;
                buildMobileAccordionsFromTable('tableResumenOperativo', 'resumenOperativoMobileContainer', 'Resumen Operativo', '');
                buildMobileAccordionsFromTable('tableVentasSegmento', 'ventasSegmentoMobileContainer', 'Segmentos de Venta', formatCurrency(kpis.ingresos));
                buildMobileAccordionsFromTable('tableCostosSegmento', 'costosSegmentoMobileContainer', 'Desglose de Costos', formatCurrency(totalCost));
                const currentOpex = (lastData.pnl && lastData.pnl.opexDetalle) ? Object.values(lastData.pnl.opexDetalle).reduce((acc, val) => acc + val, 0) : 0;
                buildMobileAccordionsFromTable('tableOpex', 'opexMobileContainer', 'Detalle de Gastos OPEX', formatCurrency(currentOpex));
            }
        }
    });
});

function exportToExcelSuite(data) {
    const wb = XLSX.utils.book_new();

    // 1. Resumen_Ejecutivo (Visual formatting)
    let totalVentas = 0;
    let totalEbitda = 0;
    data.forEach(d => {
        totalVentas += d.kpis?.ingresos || 0;
        totalEbitda += d.kpis?.ebitda || 0;
    });
    let margenPromedio = totalVentas ? (totalEbitda / totalVentas) * 100 : 0;

    const resumenData = [
        { A: "RESUMEN EJECUTIVO FINANCIERO", B: "" },
        { A: "", B: "" },
        { A: "VENTAS TOTALES", B: formatRawCurrency(totalVentas) },
        { A: "EBITDA ACUMULADO", B: formatRawCurrency(totalEbitda) },
        { A: "MARGEN PROMEDIO", B: margenPromedio.toFixed(2) + "%" },
        { A: "", B: "" },
        { A: "PERIODO ANALIZADO", B: `${data[0]?.date || ''} - ${data[data.length-1]?.date || ''}` }
    ];
    const resSheet = XLSX.utils.json_to_sheet(resumenData, { skipHeader: true });
    XLSX.utils.book_append_sheet(wb, resSheet, "RESUMEN_EJECUTIVO");

    // 2. KPI_Dashboard
    const kpiData = data.map(d => ({
        Periodo: d.date,
        Ingresos: d.kpis.ingresos,
        EBITDA: d.kpis.ebitda,
        "Margen EBITDA %": (d.kpis.margen_ebitda * 100).toFixed(2) + "%",
        "Utilidad Neta": d.pnl?.netIncome || 0,
        "Cash Flow": d.kpis.cashflow
    }));
    const kpiSheet = XLSX.utils.json_to_sheet(kpiData);
    XLSX.utils.book_append_sheet(wb, kpiSheet, "KPI_Dashboard");

    // 3. PnL_Detallado
    const pnlTable = document.getElementById('pnlDetailedTable');
    if(pnlTable) {
        const pnlSht = XLSX.utils.table_to_sheet(pnlTable, {raw: false});
        XLSX.utils.book_append_sheet(wb, pnlSht, "PnL_Detallado");
    }

    // 4. Balance_Sheet
    const balTable = document.getElementById('balanceTable');
    if(balTable) {
        const balSht = XLSX.utils.table_to_sheet(balTable, {raw: false});
        XLSX.utils.book_append_sheet(wb, balSht, "Balance_Sheet");
    }

    // 5. Cash_Flow
    const cfTable = document.getElementById('cashflowTable');
    if(cfTable) {
        const cfSht = XLSX.utils.table_to_sheet(cfTable, {raw: false});
        XLSX.utils.book_append_sheet(wb, cfSht, "Cash_Flow");
    }

    // 6. Estados_Financieros
    const estTable = document.getElementById('table-estados');
    if(estTable) {
        const estSht = XLSX.utils.table_to_sheet(estTable, {raw: false});
        XLSX.utils.book_append_sheet(wb, estSht, "Estados_Financieros");
    }

    XLSX.writeFile(wb, "Reporte_Ejecutivo_CEO.xlsx");
}

async function handleFileUpload(e) {
    const file = e.target.files[0];
    if (!file) return;
    
    const statusEl = document.getElementById('engineStatus');
    statusEl.innerHTML = "⏳ Procesando archivo...";

    loader.style.display = 'flex';
    
    const reader = new FileReader();
    reader.onload = async (event) => {
        try {
            const workbook = XLSX.read(new Uint8Array(event.target.result), { type: 'array', cellDates: true });
            const engineResult = financialEngine(workbook);
            
            if (engineResult.error || !engineResult.data || engineResult.data.length === 0) {
                const errorMsg = engineResult.error || "No se pudieron extraer datos numéricos del archivo. Asegúrese de que las filas de 'Ventas' o 'Ingresos' tengan valores numéricos.";
                showError(errorMsg);
                loader.style.display = 'none';
                return;
            }

            statusEl.innerHTML = "⏳ Validando datos para IA...";
            
            // Tomar el último periodo para el LLM
            const lastData = engineResult.data[engineResult.data.length - 1];
            
            if (!lastData || !lastData.balance) {
                showError("Estructura de datos incompleta en el archivo.");
                loader.style.display = 'none';
                return;
            }

            const pnlResult = {
                ventas: engineResult.data.map(d => d.kpis.ingresos),
                ebitda: engineResult.data.map(d => d.kpis.ebitda)
            };

            const balanceResult = {
                activos: lastData.balance.activos || 0,
                pasivos: lastData.balance.pasivos || 0,
                patrimonio: lastData.balance.patrimonio || 0
            };

            // 🧠 1. Construir input
            const llmInput = buildLLMInput({
                pnlData: pnlResult,
                balanceData: balanceResult,
                source: "excel_upload"
            });

            // 🔍 2. Validar
            const validation = validateLLMInput(llmInput);

            if (!validation.isValid) {
                console.warn("Validation Warnings:", validation.errors);
                statusEl.innerHTML = `✅ Modelo Local: ${engineResult.modelType}`;
                globalFinancialData = engineResult.data;
                renderDashboard(globalFinancialData);
            } else {
                statusEl.innerHTML = "🚀 Consultando Analista AI...";
                try {
                    const aiResponse = await callAI(llmInput);
                    statusEl.innerHTML = `✅ IA: Análisis Completado`;
                    
                    // CRITICAL FIX: Do NOT overwrite the full dataset with the single AI month
                    // Keep the engine result but maybe merge AI alerts/insights if needed
                    globalFinancialData = engineResult.data;
                    
                    // Option: Attach AI insights to the last month
                    const lastIdx = globalFinancialData.length - 1;
                    if (aiResponse.alerts) {
                        globalFinancialData[lastIdx].alerts = [...(globalFinancialData[lastIdx].alerts || []), ...aiResponse.alerts];
                    }
                    
                    renderDashboard(globalFinancialData);
                } catch (aiErr) {
                    console.error("AI Error:", aiErr);
                    statusEl.innerHTML = `⚠️ Usando motor local.`;
                    globalFinancialData = engineResult.data;
                    renderDashboard(globalFinancialData);
                }
            }
            
            loader.style.display = 'none';
        } catch (err) {
            console.error("Engine Error:", err);
            showError("Error crítico: " + err.message);
            loader.style.display = 'none';
        }
    };
    reader.readAsArrayBuffer(file);
}

async function callAI(payload) {
    const response = await ai.models.generateContent({
        model: "gemini-3-flash-preview",
        contents: `Actúa como un Senior Financial Analyst y analiza estos datos de P&L y Balance.
        
        INSTRUCCIONES:
        1. Devuelve un JSON estrictamente válido.
        2. Proporciona insights sobre el EBITDA y la eficiencia operativa.
        3. Identifica variaciones atípicas.
        
        ESTRUCTURA REQUERIDA (NO OMITIR CAMPOS):
        {
          "date": "Periodo Actual",
          "kpis": { "ingresos": number, "ebitda": number, "margen_ebitda": number, "cashflow": number },
          "balance": { "activos": number, "pasivos": number, "patrimonio": number, "cuadra": boolean },
          "pnl": { "categorias": { "Categoria": valor, ... }, "segments": {} },
          "alerts": ["string"]
        }

        DATOS PARA ANALIZAR:
        ${JSON.stringify(payload, null, 2)}`,
        config: {
            responseMimeType: "application/json"
        }
    });

    return JSON.parse(response.text);
}

function showError(msg) {
    const statusEl = document.getElementById('engineStatus');
    statusEl.style.background = '#fee2e2';
    statusEl.style.color = '#991b1b';
    statusEl.style.borderColor = '#fecaca';
    statusEl.innerHTML = `❌ ${msg}`;
}

/**
 * 🚀 MOBILE ACCORDION GENERATOR
 * Converts desktop tables into mobile-friendly vertical cards wrapped in accordions.
 */
function buildMobileAccordionsFromTable(tableId, containerId, customTitle = null, customSummary = null) {
    const table = document.getElementById(tableId);
    if (!table) return;
    const isMobile = window.innerWidth < 768;
    const container = document.getElementById(containerId);
    if(!container) return;

    if (!isMobile) {
        table.style.display = '';
        container.style.display = 'none';
        return;
    }

    // Determine if table is inside a section or just bare
    table.style.setProperty('display', 'none', 'important');
    container.style.display = 'block';

    const ths = Array.from(table.querySelectorAll('thead th'));
    const headers = ths.slice(1).map(th => th.innerText);

    const rows = Array.from(table.querySelectorAll('tbody tr'));
    
    let html = '';
    let currentGroupHtml = '';
    let currentGroupTitle = customTitle || 'Categoría / Cuentas';
    let currentGroupSummary = customSummary || '';
    
    let isSingleGroupTable = !rows.some(tr => tr.classList.contains('row-category'));

    const flushGroup = (newTitle, newSummary) => {
        if (currentGroupHtml !== '') {
             html += `<div class="mobile-accordion-group">
                <div class="mobile-accordion-header" onclick="this.nextElementSibling.classList.toggle('open')">
                    <div style="display:flex; flex-direction:column; gap:4px; max-width:85%;">
                        <span style="text-transform: uppercase;">${currentGroupTitle}</span>
                        ${currentGroupSummary ? `<span style="font-size:12px; color:var(--primary); font-weight: 800;">TOTAL: ${currentGroupSummary}</span>` : ''}
                    </div>
                    <i data-lucide="chevron-down" style="width:20px;height:20px;"></i>
                </div>
                <!-- Remove display none by default if it's a single group table so it opens by default or let user open it -->
                <div class="mobile-accordion-content ${isSingleGroupTable ? 'open' : ''}">
                    ${currentGroupHtml}
                </div>
             </div>`;
        }
        currentGroupHtml = '';
        currentGroupTitle = newTitle || customTitle || 'Categoría';
        currentGroupSummary = newSummary || '';
    };

    rows.forEach((tr, i) => {
        const tds = Array.from(tr.querySelectorAll('td'));
        if (tds.length < 2) return; // empty row or spacer

        const label = tds[0].innerText;
        const vals = tds.slice(1).map(td => td.innerText);
        
        const isTotal = tr.classList.contains('row-total');
        const isCategory = tr.classList.contains('row-category');
        
        if (isCategory) {
            flushGroup(label);
        } else if (isTotal) {
            // Find a valid numerical string to show as summary for the accordion
            let summaryVal = '';
            for(let j = vals.length - 1; j >= 0; j--) {
                if(vals[j] && vals[j] !== '-') { summaryVal = vals[j]; break; }
            }
            if(!currentGroupSummary) currentGroupSummary = summaryVal || vals[vals.length - 1];
            currentGroupHtml += createMobileCard(label, headers, vals);
            if (currentGroupTitle === 'Categoría') currentGroupTitle = label;
            
            // Only flush if we're dealing with a multi-category table like P&L
            if (!isSingleGroupTable && i < rows.length - 1) {
                flushGroup();
            }
        } else {
            currentGroupHtml += createMobileCard(label, headers, vals);
            // If it's the last row and a single group table, and we don't have a summary, we can try to guess it.
        }
    });

    flushGroup(); // flush remaining
    
    if (html === '') {
       container.innerHTML = '<div style="padding:20px; text-align:center; font-size:12px; color:var(--text-secondary);">No hay datos formatados para mostrar.</div>';
    } else {
       // Add Swipe Indicator (as requested by user)
       container.innerHTML = `<div class="swipe-indicator"> <i data-lucide="chevrons-down" style="width:14px;height:14px;display:inline-block;vertical-align:middle;"></i> Toca para interactuar</div>` + html;
       if (typeof lucide !== 'undefined') lucide.createIcons();
    }
}

function createMobileCard(label, headers, vals) {
    let cardHtml = `<div class="mobile-vertical-card">
        <div class="mobile-vertical-card-title">
            <span style="max-width:80%; overflow:hidden; text-overflow:ellipsis; white-space:nowrap;">${label}</span>
            <i data-lucide="trending-up" style="width:14px;height:14px;opacity:0.5;"></i>
        </div>`;
    vals.forEach((v, idx) => {
        if (headers[idx]) {
            cardHtml += `<div class="mobile-vertical-card-row">
                <span style="color:var(--text-secondary);">${headers[idx]}</span>
                <span style="font-weight:700;">${v}</span>
            </div>`;
        }
    });
    cardHtml += `</div>`;
    return cardHtml;
}

// Global UI Updater Function
function renderDashboard(data) {
    if (!data || data.length === 0) return;
    
    // Filtro: No permitir seleccionar datos del 2025 en el dropdown global
    const filteredForSelector = data.map((d, i) => ({ d, i })).filter(item => !isYear2025(item.d));
    
    monthSelector.innerHTML = filteredForSelector.map(item => `<option value="${item.i}">${item.d.date || 'Periodo'}</option>`).join('');
    monthSelector.style.display = 'block';
    
    // Show search input if one of the detailed views is active
    const searchWrapper = document.getElementById('searchContainerWrapper');
    if (searchWrapper) {
        const activeMenu = document.querySelector('.menu-item a.active');
        const viewsWithSearch = ['menu-resumen', 'menu-pnl', 'menu-balance', 'menu-cashflow', 'menu-estados'];
        if (activeMenu && viewsWithSearch.includes(activeMenu.id)) {
            searchWrapper.style.display = 'flex';
        }
    }
    
    const lastIdx = filteredForSelector.length > 0 ? filteredForSelector[filteredForSelector.length - 1].i : data.length - 1;
    monthSelector.value = lastIdx;
    updateUI(data, lastIdx);
}

function updateUI(data, index) {
    if (!data || !data[index]) return;
    const curr = data[index];
    
    // Identificar el anterior operativo (excluyendo el año base 2025 para comparaciones MoM)
    const operationalData = data.filter(d => !isYear2025(d));
    const currIdxInOp = operationalData.findIndex(d => d.date === curr.date);
    const prev = currIdxInOp > 0 ? operationalData[currIdxInOp - 1] : curr;
    
    // Safety guards for kpis
    const kpis = curr.kpis || { ingresos: 0, ebitda: 0, cashflow: 0, margen_ebitda: 0 };
    const prevKpis = prev.kpis || kpis;

    // Integrity Badge logic
    const integrityBadge = document.getElementById('integrityBadge');
    if (integrityBadge && curr.integrity) {
        integrityBadge.style.display = 'flex';
        if (curr.integrity.isBroken) {
            integrityBadge.className = 'integrity-fail';
            integrityBadge.innerHTML = `⚠️ Ajuste Detectado (Abs: ${formatCurrency(curr.integrity.gap)})`;
            integrityBadge.title = "La suma de Ingresos - Costos - Gastos no coincide con el EBITDA reportado por un margen > 1%";
        } else {
            integrityBadge.className = 'integrity-ok';
            integrityBadge.innerHTML = `✓ P&L Cuadrado`;
            integrityBadge.title = "Integridad de datos verificada operativamente";
        }
    }

    document.getElementById('kpi-ventas').textContent = formatCurrency(kpis.ingresos);
    document.getElementById('kpi-ebitda').textContent = formatCurrency(kpis.ebitda);

    // Safety guards for pnl categories
    const categories = (curr.pnl && curr.pnl.categorias) ? curr.pnl.categorias : {};
    const prevCategories = (prev.pnl && prev.pnl.categorias) ? prev.pnl.categorias : categories;

    const totalCost = categories["Costo de Ventas"] || 0;
    const prevTotalCost = prevCategories["Costo de Ventas"] || 0;

    document.getElementById('val-ratio').textContent = formatCurrency(totalCost);

    // Renderizar Segmentos
    const segmentsSection = document.getElementById('segments-section');
    const segmentsBody = document.getElementById('segmentsBody');
    const segments = (curr.pnl && curr.pnl.segments) ? curr.pnl.segments : {};
    const prevSegments = (prev.pnl && prev.pnl.segments) ? prev.pnl.segments : segments;
    
    if (Object.keys(segments).length > 0) {
        segmentsSection.style.display = 'block';
        segmentsBody.innerHTML = Object.entries(segments).map(([name, data]) => {
            const ventas = data.ventas || 0;
            const prevVentas = prevSegments[name] ? prevSegments[name].ventas : 0;
            const diff = ventas - prevVentas;
            const pct = kpis.ingresos !== 0 ? (ventas / kpis.ingresos) * 100 : 0;
            const color = diff >= 0 ? 'var(--success)' : 'var(--danger)'; 
            const valColor = ventas < 0 ? 'var(--danger)' : 'inherit';
            const prevColor = prevVentas < 0 ? 'var(--danger)' : 'inherit';

            return `<tr>
                <td style="font-weight:600">
                    ${name}
                    <div class="bar-container"><div class="bar-fill" style="width: ${Math.min(100, Math.max(0, pct))}%"></div></div>
                </td>
                <td style="color:${prevColor}">${formatCurrency(prevVentas)}</td>
                <td style="color:${valColor}">${formatCurrency(ventas)}</td>
                <td style="color:${color}">${diff >= 0 ? '+' : ''}${formatCurrency(diff)}</td>
                <td style="font-weight:700">${pct.toFixed(1)}%</td>
            </tr>`;
        }).join('');
    } else {
        segmentsSection.style.display = 'none';
    }

    // Renderizar Costos por Segmento (Nuevo)
    const costSegmentsSection = document.getElementById('cost-segments-section');
    const costSegmentsBody = document.getElementById('costSegmentsBody');
    if (Object.keys(segments).length > 0) {
        costSegmentsSection.style.display = 'block';
        costSegmentsBody.innerHTML = Object.entries(segments).map(([name, data]) => {
            const costos = data.costos || 0;
            const prevCostos = prevSegments[name] ? prevSegments[name].costos : 0;
            
            const diff = costos - prevCostos;
            const pctVar = prevCostos !== 0 ? (diff / Math.abs(prevCostos)) * 100 : 0;
            
            // Regla solicitada: Positivo = Verde, Negativo = Rojo
            const color = diff >= 0 ? 'var(--success)' : 'var(--danger)';
            const valColor = costos < 0 ? 'var(--danger)' : 'inherit';
            const prevColor = prevCostos < 0 ? 'var(--danger)' : 'inherit';

            return `<tr>
                <td style="font-weight:600">${name}</td>
                <td style="color:${prevColor}">${formatCurrency(prevCostos)}</td>
                <td style="color:${valColor}">${formatCurrency(costos)}</td>
                <td style="color:${color}">${diff > 0 ? '+' : ''}${formatCurrency(diff)}</td>
                <td style="color:${color}">${pctVar > 0 ? '+' : ''}${pctVar.toFixed(1)}%</td>
            </tr>`;
        }).join('');
    } else {
        costSegmentsSection.style.display = 'none';
    }

    // Renderizar Detalle OPEX
    const opexSection = document.getElementById('opex-section');
    const opexBody = document.getElementById('opexBody');
    const opexDetalle = (curr.pnl && curr.pnl.opexDetalle) ? curr.pnl.opexDetalle : {};
    const prevOpexDetalle = (prev.pnl && prev.pnl.opexDetalle) ? prev.pnl.opexDetalle : opexDetalle;

    if (Object.keys(opexDetalle).length > 0) {
        opexSection.style.display = 'block';
        opexBody.innerHTML = Object.entries(opexDetalle).map(([cat, val]) => {
            const prevVal = prevOpexDetalle[cat] || 0;
            // Unificamos lógica: val - prevVal es el impacto en la salud financiera
            // Si el monto es negativo (gasto), un incremento hacia cero es positivo
            // Si el monto es positivo (ingreso), un incremento es positivo
            const diff = val - prevVal; 
            const pct = prevVal !== 0 ? (diff / Math.abs(prevVal)) * 100 : 0;
            const color = diff >= 0 ? 'var(--success)' : 'var(--danger)'; 
            const valColor = val < 0 ? 'var(--danger)' : 'inherit';
            const prevColor = prevVal < 0 ? 'var(--danger)' : 'inherit';

            return `<tr>
                <td style="font-weight:600">${cat}</td>
                <td style="color:${prevColor}">${formatCurrency(prevVal)}</td>
                <td style="color:${valColor}">${formatCurrency(val)}</td>
                <td style="color:${color}">${diff > 0 ? '+' : ''}${formatCurrency(diff)}</td>
                <td style="color:${color}">${pct > 0 ? '+' : ''}${pct.toFixed(1)}%</td>
            </tr>`;
        }).join('');
    } else {
        opexSection.style.display = 'none';
    }

    // Renderizar Tabla Detallada
    const tableBody = document.getElementById('tableBody');
    if (Object.keys(categories).length > 0) {
        // Excluimos OPEX y Utilidad Neta para dejar solo indicadores operativos puros
        const filteredEntries = Object.entries(categories).filter(([cat]) => 
            !cat.toLowerCase().includes("opex") && 
            !cat.toLowerCase().includes("extraordinarios") &&
            !cat.toLowerCase().includes("utilidad")
        );

        tableBody.innerHTML = filteredEntries.map(([cat, val]) => {
            const prevVal = prevCategories[cat] || 0;
            
            const diff = val - prevVal;
            const pct = prevVal !== 0 ? (diff / Math.abs(prevVal)) * 100 : 0;
            
            // Unificamos lógica de color con el resto de tablas del resumen (Positivo = Verde, Negativo = Rojo)
            // Para indicadores operativos, un incremento suele ser positivo
            const color = diff >= 0 ? 'var(--success)' : 'var(--danger)';

            const valColor = val < 0 ? 'var(--danger)' : 'inherit';
            const prevColor = prevVal < 0 ? 'var(--danger)' : 'inherit';
            
            return `<tr>
                <td style="font-weight:600">${cat}</td>
                <td style="color:${prevColor}">${formatCurrency(prevVal)}</td>
                <td style="color:${valColor}">${formatCurrency(val)}</td>
                <td style="color:${color}">${diff > 0 ? '+' : ''}${formatCurrency(diff)}</td>
                <td style="color:${color}">${pct > 0 ? '+' : ''}${pct.toFixed(1)}%</td>
            </tr>`;
        }).join('');
    }

    const statusEl = document.getElementById('engineStatus');
    if (statusEl && curr.pnl && curr.pnl.detectedRows) {
        statusEl.innerHTML = `✅ Datos Detectados:<br>
            <b>Ingresos:</b> "${curr.pnl.detectedRows.ingresos || '?'}"<br>
            <b>EBITDA:</b> "${curr.pnl.detectedRows.ebitda || '?'}"<br>
            <b>OPEX:</b> "${curr.pnl.detectedRows.opex || '?'}"<br>
            <b>Balance:</b> "${curr.pnl.detectedRows.activos || 'No detectado'}"`;
    }
    
    document.getElementById('periodLabel').textContent = `Periodo de Análisis: ${curr.date || 'Actual'}`;
    updateTrend('sub-ventas', kpis.ingresos, prevKpis.ingresos);
    
    // EBITDA Trend + Margin
    const margin = ((kpis.margen_ebitda || 0) * 100).toFixed(1);
    updateTrend('sub-ebitda', kpis.ebitda, prevKpis.ebitda, ` | Margen: ${margin}%`);
    
    // Costos de Ventas Trend
    updateTrend('sub-ratio', totalCost, prevTotalCost);

    // Render Detailed P&L (Passing current selected index for rolling window)
    renderDetailedPnL(data, index);
    
    // Render Balance Sheet
    renderBalanceSheet(data, index);

    // Render Cash Flow
    renderCashFlow(data, index);

    // 🚀 NEW: Render KPI Dashboard
    renderKPIDashboard(data, index);

    // 🚀 NEW: Render Estados Financieros
    renderEstadosFinancieros(data, index);

    // Build Mobile Views
    setTimeout(() => {
        buildMobileAccordionsFromTable('pnlDetailedTable', 'pnlMobileContainer');
        buildMobileAccordionsFromTable('balanceTable', 'balanceMobileContainer');
        buildMobileAccordionsFromTable('covenantTable', 'covenantMobileContainer');
        buildMobileAccordionsFromTable('cashflowTable', 'cashflowMobileContainer');
        buildMobileAccordionsFromTable('cfMetricsTable', 'cfMetricsMobileContainer');
        buildMobileAccordionsFromTable('table-estados', 'estadosMobileContainer');
        
        // Resumen Header Acccords
        buildMobileAccordionsFromTable('tableResumenOperativo', 'resumenOperativoMobileContainer', 'Resumen Operativo', '');
        buildMobileAccordionsFromTable('tableVentasSegmento', 'ventasSegmentoMobileContainer', 'Segmentos de Venta', formatCurrency(kpis.ingresos));
        buildMobileAccordionsFromTable('tableCostosSegmento', 'costosSegmentoMobileContainer', 'Desglose de Costos', formatCurrency(totalCost));
        
        const currOpex = (curr.pnl && curr.pnl.opexDetalle) ? Object.values(curr.pnl.opexDetalle).reduce((acc, val) => acc + val, 0) : 0;
        buildMobileAccordionsFromTable('tableOpex', 'opexMobileContainer', 'Detalle de Gastos OPEX', formatCurrency(currOpex));
        
        // Trigger account search filter if active
        const searchInput = document.getElementById('accountSearch');
        if (searchInput && searchInput.value.trim() !== '') {
            searchInput.dispatchEvent(new Event('input'));
        }
    }, 50);
}

/**
 * Helper to identify periods from 2025
 */
function isYear2025(d) {
    if (!d) return false;
    const dt = d.sortDate;
    const normDate = normalizeText(d.date || "");
    
    if (dt && typeof dt.getFullYear === 'function' && dt.getFullYear() === 2025) return true;
    if (dt && typeof dt === 'string') {
        const dObj = new Date(dt);
        if (!isNaN(dObj) && dObj.getFullYear() === 2025) return true;
    }
    if (normDate.includes("2025") || normDate.includes("-25") || normDate.includes("/25") || normDate.includes(" 25")) return true;
    
    return false;
}

/**
 * Render the Balance Sheet Table
 */
function renderBalanceSheet(data, selectedIndex = -1) {
    const headerEl = document.getElementById('balanceHeader');
    const bodyEl = document.getElementById('balanceBody');
    const periodLabel = document.getElementById('balancePeriodLabel');
    if (!headerEl || !bodyEl || !data || data.length === 0) return;

    const endIdx = selectedIndex >= 0 ? selectedIndex : data.length - 1;
    const curr = data[endIdx];
    
    const startIdx = Math.max(0, endIdx - 5);
    let visibleMonths = data.slice(startIdx, endIdx + 1);

    // Ocultar 2025 (Solo usado para YOY)
    visibleMonths = visibleMonths.filter(m => !isYear2025(m));
    
    const periods = visibleMonths.map(d => d.date);

    periodLabel.textContent = `Situación Financiera al ${curr.date}`;

    headerEl.innerHTML = `
        <tr>
            <th>Concepto / Cuenta de Balance</th>
            ${periods.map(p => `<th>${p}</th>`).join('')}
        </tr>
    `;

    // Extract concepts
    let allConcepts = [];
    visibleMonths.forEach(d => {
        if (d.balance && d.balance.fullRows) {
            d.balance.fullRows.forEach(row => {
                if (!allConcepts.includes(row.concept)) allConcepts.push(row.concept);
            });
        }
    });

    // 1. Clasificación: Balance vs Covenants
    let balanceConcepts = [];
    let covenantConcepts = [];
    
    allConcepts.forEach(c => {
        const n = normalizeText(c);
        
        // Filtros solicitados por usuario (REA, redundantes y filas técnicas)
        if (n === "covenant deuda" || n === "rea" || n.trim() === "" || 
            n === "pasivos excluye deuda subordinada" || 
            n === "patrimonio incluye deuda subordinada" || 
            n === "pasivos - deuda subordinada") return;

        const isDebtRow = n.includes("deuda bruta") || n.includes("deuda total") || n.includes("deuda subordinada") || 
                          n.includes("deuda sin subordinada") || n.includes("deuda neta") || n.includes("ebitda");
        const isRatioRow = n.includes("apalancamiento") || n.includes("capacidad") || n.includes("razon corriente") || n.includes("covenant");
        
        // Efectivo es parte del bloque si está rodeado de deuda
        const isCovenant = isDebtRow || isRatioRow || n === "efectivo" || n.includes("ebitda r12");
        
        if (isCovenant) covenantConcepts.push(c);
        // Efectivo debe estar en AMBOS (Covenant y Balance)
        if (!isCovenant || n === "efectivo") balanceConcepts.push(c);
    });

    // Ordenamiento Estricto según imagen del Excel (Linear Extraction Mode)
    const getCovenantRank = (concept) => {
        const n = normalizeText(concept);
        if (n === "deuda bruta") return 1;
        if (n === "efectivo") return 2;
        if (n === "deuda neta" && !n.includes("ebitda") && !n.includes("subordinada")) return 3;
        if (n.includes("ltm ebitda") || (n.includes("ebitda") && !n.includes("ratio") && !n.includes("r12"))) return 4;
        
        if (n === "deuda total") return 5;
        if (n === "deuda subordinada") return 6;
        if (n === "deuda sin subordinada") return 7;
        if (n === "deuda neta sin subordinada") return 9;

        if (n.includes("deuda neta/ebitda") || n.includes("r12")) return 10;
        if (n.includes("apalancamiento")) return 11;
        if (n.includes("capacidad")) return 12;
        if (n.includes("razon corriente") || n.includes("corriente")) return 13;
        return 100;
    };

    covenantConcepts.sort((a, b) => getCovenantRank(a) - getCovenantRank(b));

    // 2. Filtro y Reordenamiento para el Balance
    let filteredBalance = [];
    let isSkipping = false;
    let foundGrandTotal = false;
    
    // Identificar posiciones especiales
    const utilidadesRetenidasIdx = balanceConcepts.findIndex(c => normalizeText(c).includes("utilidades retenidas"));
    const beneficioNetoConcept = balanceConcepts.find(c => {
        const n = normalizeText(c);
        return (n.includes("beneficio neto") && !n.includes("utilidades")) || 
               (n.includes("utilidad del ejercicio") && !n.includes("retenidas")) || 
               (n.includes("ganancia del periodo")) ||
               (n.includes("resultado del ejercicio") && !n.includes("retenidas"));
    });
    
    balanceConcepts.forEach(concept => {
        if (foundGrandTotal) return;
        const norm = normalizeText(concept);
        
        // Regla de Parada 1: Eliminar redundantes y cabeceras de Excel
        if (norm === "total pasivo y patrimonio" || norm === "total pasivo y capital") return;
        if (norm === "concepto" || norm === "cuentas" || norm === "descripcion" || norm === "balance sheet" || norm === "detalle") return;
        
        // Exclusión agresiva de Activos/Pasivos/Patrimonio como cabeceras puras
        const isHeaderOnly = norm === "activos" || norm === "pasivos" || norm === "patrimonio" || 
                             norm === "capital" || norm === "pasivo y capital" || 
                             norm === "activo" || norm === "pasivo" ||
                             norm.startsWith("activos:") || norm.startsWith("pasivos:") || 
                             norm.startsWith("patrimonio:");
        
        if (isHeaderOnly) return;
        if (norm.includes("estado de situacion") || norm.includes("reporte pa") || norm.includes("en mdop")) return;

        // Regla de Parada Final Fuerte
        if (norm.includes("total pasivo") && (norm.includes("capital") || norm.includes("accionista"))) {
            if (concept.trim().length > 10) { // Evitar falsos positivos cortos
                filteredBalance.push(concept);
                foundGrandTotal = true;
                return;
            }
        }

        // Evitar duplicar beneficio neto (se insertará debajo de utilidades retenidas)
        if (concept === beneficioNetoConcept && utilidadesRetenidasIdx !== -1) return;

        // Limpieza de firmas
        const isPatrimonioItem = norm.includes("utilidad") || norm.includes("beneficio") || norm.includes("ganancia") || 
                                 norm.includes("reserva") || norm.includes("capital") || norm.includes("patrimonio") || 
                                 norm.includes("rea") || norm.includes("resultados acumulados") || 
                                 norm.includes("ajuste") || norm.includes("manos de terceros");

        if (isSkipping && isPatrimonioItem) isSkipping = false;

        if (!isSkipping || isPatrimonioItem) {
            filteredBalance.push(concept);
            // Inserción de Beneficio Neto debajo de Utilidades Retenidas
            if (concept === balanceConcepts[utilidadesRetenidasIdx] && beneficioNetoConcept) {
                if (!filteredBalance.some(c => c === beneficioNetoConcept)) {
                    filteredBalance.push(beneficioNetoConcept);
                }
            }
        }
    });

    // Asegurar que "Efectivo" sea la primera fila del Balance General Consolidado
    const cashIndex = filteredBalance.findIndex(c => normalizeText(c) === "efectivo");
    if (cashIndex > 0) {
        const cashRow = filteredBalance.splice(cashIndex, 1)[0];
        filteredBalance.unshift(cashRow);
    }

    const renderRows = (concepts, targetBodyId) => {
        const bodyEl = document.getElementById(targetBodyId);
        if (!bodyEl) return;
        
        bodyEl.innerHTML = concepts.map(concept => {
            const norm = normalizeText(concept);
            const labelLower = norm;
            const isTotal = labelLower.includes("total") || labelLower.includes("sumas") || 
                            labelLower.includes("activo") || labelLower.includes("pasivo") || 
                            labelLower.includes("patrimonio") ||
                            labelLower.includes("ebitda") || labelLower.includes("apalancamiento") || 
                            labelLower.includes("capacidad de pago") || labelLower.includes("razon corriente");
            const isSubRow = concept.startsWith("  ") || concept.startsWith("\t") || 
                             norm.includes("acumulado") || norm.includes("depreciacion") || norm.includes("impuestos") ||
                             norm.includes("ganancia acumulada") || 
                             norm.includes("beneficio neto") || norm.includes("ganancia del periodo") ||
                             norm.includes("resultado del ejercicio") ||
                             norm.includes("activo en manos de terceros");
            
            const isMainCategory = (concept.trim() === concept.trim().toUpperCase() || 
                                    labelLower.includes("activos") || labelLower.includes("pasivos") || 
                                    labelLower.includes("patrimonio") || norm.includes("covenant") ||
                                    norm.includes("utilidades retenidas") || norm.includes("revaluacion de activos")) 
                                    && !isTotal && concept.trim().length > 3;

            const periodCells = visibleMonths.map(period => {
                const row = period.balance?.fullRows?.find(r => r.concept === concept);
                let val = row ? row.values[period.date] || 0 : 0;

                // Fallback para Beneficio Neto: si es 0 en el balance, tomarlo del P&L
                if (norm.includes("beneficio neto") || norm.includes("ganancia del periodo") || 
                    norm.includes("utilidad del ejercicio") || norm.includes("resultado del ejercicio") || 
                    norm.includes("utilidad neta") || norm.includes("ganancia neta")) {
                    if (val === 0 && period.pnl?.netIncome) {
                        val = period.pnl.netIncome;
                    }
                }
                
                const isRatio = (norm.includes("ratio") || norm.includes("indice") || norm.includes("razon") || 
                                 norm.includes("apalancamiento") || norm.includes("capacidad") || 
                                 norm.includes("ebitda r12") || norm.includes("ebitda ltm") ||
                                 norm.includes("deuda neta/ebitda") || concept.includes(" x ") || concept.endsWith(" x")) && 
                                 !norm.includes("cxp") && !norm.includes("otras cxp") && !norm.includes("cxc") && !norm.includes("pagar") && !norm.includes("cobrar");
                
                const color = val < 0 ? 'var(--danger)' : 'inherit';
                let displayVal;
                
                if (isRatio) {
                    displayVal = (val !== 0) ? (typeof val === 'number' ? val.toFixed(2) : val) + 'x' : "-";
                } else if ((norm.includes("covenant") || norm.includes("apalancamiento") || norm.includes("capacidad") || norm.includes("razon corriente") || norm.includes("ebitda r12")) && val !== 0 && !norm.includes("mdo") && !norm.includes("pagar") && !norm.includes("cobrar")) {
                    displayVal = (typeof val === 'number' ? val.toFixed(2) : val) + 'x';
                } else {
                    displayVal = formatCurrency(val);
                }

                // Si es una categoría principal y el valor es 0, ocultamos el valor para evitar confusión
                if (isMainCategory && val === 0) displayVal = "";

                return `<td style="color:${color}">${displayVal}</td>`;
            }).join('');

            let displayLabel = concept;
            if (norm === "ganancia del periodo") displayLabel = "Beneficio Neto del Periodo";
            
            let rowClass = isTotal ? 'row-total' : '';
            if (isMainCategory && !isSubRow) rowClass = 'row-category';
            
            const cellClass = isSubRow ? 'row-indent' : '';

            return `<tr class="${rowClass}">
                <td class="${cellClass}">${displayLabel}</td>
                ${periodCells}
            </tr>`;
        }).join('');
    };

    if (filteredBalance.length === 0) {
        bodyEl.innerHTML = `<tr><td colspan="${periods.length + 1}" style="text-align:center; padding:40px;">No se encontraron filas detalladas de Balance.</td></tr>`;
    } else {
        renderRows(filteredBalance, 'balanceBody');
    }
    
    // Render Covenant Section
    if (covenantConcepts.length > 0) {
        document.getElementById('covenant-section').style.display = 'block';
        document.getElementById('covenantHeader').innerHTML = `
            <tr>
                <th>Concepto / Ratio de Deuda</th>
                ${periods.map(p => `<th>${p}</th>`).join('')}
            </tr>
        `;
        renderRows(covenantConcepts, 'covenantBody');
    } else {
        document.getElementById('covenant-section').style.display = 'none';
    }
}

/**
 * Render the Cash Flow Table
 */
function renderCashFlow(data, selectedIndex = -1) {
    const headerEl = document.getElementById('cashflowHeader');
    const bodyEl = document.getElementById('cashflowBody');
    const periodLabel = document.getElementById('cashflowPeriodLabel');
    if (!headerEl || !bodyEl || !data || data.length === 0) return;

    const endIdx = selectedIndex >= 0 ? selectedIndex : data.length - 1;
    const curr = data[endIdx];
    
    const startIdx = Math.max(0, endIdx - 5);
    let visibleMonths = data.slice(startIdx, endIdx + 1);

    // Ocultar 2025 (Solo usado para YOY)
    visibleMonths = visibleMonths.filter(m => !isYear2025(m));
    
    const periods = visibleMonths.map(d => d.date);
    periodLabel.textContent = `Análisis de Ciclo de Caja al ${curr.date}`;

    headerEl.innerHTML = `
        <tr>
            <th>Concepto / Flujo de Efectivo</th>
            ${periods.map(p => `<th>${p}</th>`).join('')}
        </tr>
    `;

    // Definition of rows in order
    const rowSpec = [
        { key: 'beginning', label: 'Efectivo Inicial', isHeader: true },
        { key: 'ebitda', label: 'EBITDA (Origen P&L)', isSource: true },
        { type: 'separator', label: 'Cambios en Capital de Trabajo' },
        { key: 'cxc', label: ' (Aumento)/Disminución CxC', indent: true },
        { key: 'inv', label: ' (Aumento)/Disminución Inventario', indent: true },
        { key: 'cxp', label: ' Aumento/(Disminución) CxP', indent: true },
        { key: 'wc', label: 'Total Cambios Capital Trabajo', isTotal: true },
        { type: 'separator', label: 'Otros Ajustes Operativos' },
        { key: 'taxes', label: 'Impuestos Pagados', indent: true },
        { key: 'extraordinary', label: 'Gastos/Ingresos Extraordinarios', indent: true },
        { key: 'operating', label: 'FLUJO DE CAJA OPERATIVO', isTotal: true },
        { type: 'separator', label: 'Actividades de Inversión' },
        { key: 'capex', label: 'Inversiones de Capital (CAPEX)', indent: true },
        { type: 'separator', label: 'Actividades de Financiamiento' },
        { key: 'netDebt', label: 'Variación Deuda Neta', indent: true },
        { key: 'interest', label: 'Gastos Financieros / Intereses', indent: true },
        { key: 'dividends', label: 'Actividades con Accionistas / Otros', indent: true },
        { key: 'financing', label: 'Total Flujo Financiamiento', isTotal: true },
        { key: 'change', label: 'VARIACIÓN NETA DE EFECTIVO', isHeader: true },
        { key: 'ending', label: 'Efectivo Final', isHeader: true }
    ];

    bodyEl.innerHTML = rowSpec.map(spec => {
        if (spec.type === 'separator') {
            return `<tr class="row-category"><td colspan="${periods.length + 1}" style="background:rgba(0,0,0,0.02); font-weight:700; font-size:0.75rem; color:var(--text-secondary); text-transform:uppercase; letter-spacing:0.5px;">${spec.label}</td></tr>`;
        }

        const cells = visibleMonths.map(period => {
            let val = 0;
            if (spec.key === 'ebitda') {
                val = period.kpis?.ebitda || 0;
            } else {
                val = period.cashflowDetail?.[spec.key] || 0;
            }

            const color = val < 0 ? 'var(--danger)' : 'inherit';
            return `<td style="color:${color}">${formatCurrency(val)}</td>`;
        }).join('');

        let className = spec.isTotal ? 'row-total' : (spec.isHeader ? 'row-category' : '');
        let cellClass = spec.indent ? 'row-indent' : '';
        if (spec.isHeader) className += ' row-highlight';

        return `<tr class="${className}">
            <td class="${cellClass}">${spec.label}</td>
            ${cells}
        </tr>`;
    }).join('');

    // Metrics Section (DSO, DPO, DIO)
    const metricsBody = document.getElementById('cfMetricsBody');
    const metricsHeader = document.getElementById('cfMetricsHeader');
    const metricsSection = document.getElementById('cf-metrics-section');

    const hasMetrics = visibleMonths.some(m => m.cashflowDetail?.dso || m.cashflowDetail?.dpo || m.cashflowDetail?.dio);
    
    if (hasMetrics && metricsBody && metricsHeader) {
        metricsSection.style.display = 'block';
        metricsHeader.innerHTML = `<tr><th>Indice de Eficiencia (Días)</th>${periods.map(p => `<th>${p}</th>`).join('')}</tr>`;
        
        const metricRows = [
            { key: 'dso', label: 'DSO (Días Cuentas por Cobrar)' },
            { key: 'dpo', label: 'DPO (Días Cuentas por Pagar)' },
            { key: 'dio', label: 'DIO (Días Rotación Inventario)' }
        ];

        metricsBody.innerHTML = metricRows.map(m => {
            const cells = visibleMonths.map(p => {
                const val = p.cashflowDetail?.[m.key] || 0;
                return `<td>${Math.round(val)} días</td>`;
            }).join('');
            return `<tr><td>${m.label}</td>${cells}</tr>`;
        }).join('');
    } else {
        metricsSection.style.display = 'none';
    }
}

/**
 * Render the full P&L Detail Table with a 6-month rolling window
 */
function renderDetailedPnL(data, selectedIndex = -1) {
    const headerEl = document.getElementById('pnlDetailedHeader');
    const bodyEl = document.getElementById('pnlDetailedBody');
    if (!headerEl || !bodyEl || !data || data.length === 0) return;

    // Use current selection as anchor, or last month if not specified
    const endIdx = selectedIndex >= 0 ? selectedIndex : data.length - 1;
    // Show up to 6 months including the selected one (min 1)
    const startIdx = Math.max(0, endIdx - 5);
    
    // Filtro: No mostrar 2025 en el P&L
    const visibleMonths = data.slice(startIdx, endIdx + 1).filter(d => !isYear2025(d));
    const periods = visibleMonths.map(d => d.date);
    
    // Header
    headerEl.innerHTML = `
        <tr>
            <th>Concepto / Cuenta</th>
            ${periods.map(p => `<th>${p}</th>`).join('')}
            <th>Acum. Periodo</th>
        </tr>
    `;

    // Extract all unique account concepts from visible periods
    let allConcepts = [];
    visibleMonths.forEach(d => {
        if (d.pnl && d.pnl.fullRows) {
            d.pnl.fullRows.forEach(row => {
                if (!allConcepts.includes(row.concept)) allConcepts.push(row.concept);
            });
        }
    });

    // Filtros solicitados:
    allConcepts = allConcepts.filter(c => {
        const nc = normalizeText(c);
        if (nc === "concepto" || nc === "cuentas" || nc === "descripcion" || nc === "p&l" || nc === "resultado" || nc === "detalle") return false;
        if (nc.includes("en mdop") || nc.includes("reporte pa") || nc.includes("seguimiento gerencial") || nc.includes("margen operacional")) return false;
        return nc !== "eb usd" && nc !== "eb usd$" && nc !== "margen neto";
    });

    // 2. Eliminar desde "PPE acumulado" hacia abajo
    const ppeIndex = allConcepts.findIndex(c => normalizeText(c).includes("ppe acumulado"));
    if (ppeIndex !== -1) {
        allConcepts = allConcepts.slice(0, ppeIndex);
    }

    if (allConcepts.length === 0) {
        bodyEl.innerHTML = `<tr><td colspan="${periods.length + 2}" style="text-align:center; padding:40px;">No se encontraron filas detalladas en el reporte para este periodo.</td></tr>`;
        return;
    }

    bodyEl.innerHTML = allConcepts.map(concept => {
        let periodAccumTotal = 0;
        const normConcept = normalizeText(concept);
        const isPercentage = normConcept.includes("margen neto") || normConcept.includes("margen ebitda") || normConcept.includes("margen bruto ordinario") || normConcept.includes("margen operacional");
        const isFX = normConcept === "fx" || normConcept.includes("tasa");

        // Calculate YTD (Year to Date) total from the first data point up to endIdx
        for (let k = 0; k <= endIdx; k++) {
            const periodData = data[k];
            if (isYear2025(periodData)) continue; // 🚨 No acumular el año base en el YTD del P&L
            
            const row = periodData.pnl?.fullRows?.find(r => r.concept === concept);
            const val = row ? row.values[periodData.date] || 0 : 0;
            periodAccumTotal += val;
        }

        const periodCells = visibleMonths.map(period => {
            const row = period.pnl?.fullRows?.find(r => r.concept === concept);
            const val = row ? row.values[period.date] || 0 : 0;
            
            const color = val < 0 ? 'var(--danger)' : 'inherit';
            
            let displayVal;
            if (isPercentage) {
                displayVal = formatPercent(val);
            } else if (isFX) {
                displayVal = val.toFixed(2);
            } else {
                displayVal = formatCurrency(val);
            }

            return `<td style="color:${color}">${displayVal}</td>`;
        }).join('');

        const labelLower = concept.toLowerCase();
        const isTotal = labelLower.includes("total") || 
                        labelLower.includes("ebitda") || 
                        labelLower.includes("utilidad") ||
                        labelLower.includes("resultado") ||
                        labelLower.includes("ggadm") ||
                        labelLower.includes("ventas netas") ||
                        labelLower.includes("costo de venta") ||
                        labelLower.includes("ebit");
        
        const isSubRow = concept.startsWith("  ") || concept.startsWith("\t") || 
                         concept.toLowerCase().includes("costos ") || 
                         concept.toLowerCase().includes("gastos ") || 
                         concept.toLowerCase().includes("impuestos") || 
                         normConcept.includes("diferencial cambiario") || 
                         normConcept.includes("ingresos financieros") || 
                         normConcept.includes("gastos extraordinarios");
        const rowClass = isTotal ? 'row-total' : '';
        const cellClass = isSubRow ? 'row-indent' : '';

        // Acumulado del periodo (no tiene mucho sentido para FX, pero lo dejamos por consistencia de tabla)
        let displayAccum;
        if (isPercentage) {
            // El acumulado de un porcentaje es raramente la suma, pero mostramos el valor del último mes o promedio?
            // Por simplicidad y evitar confusiones, si es porcentaje no mostramos acumulado o mostramos el del periodo actual
            const lastVal = visibleMonths[visibleMonths.length - 1].pnl?.fullRows?.find(r => r.concept === concept)?.values[visibleMonths[visibleMonths.length - 1].date] || 0;
            displayAccum = formatPercent(lastVal);
        } else if (isFX) {
            displayAccum = "-";
        } else {
            displayAccum = formatCurrency(periodAccumTotal);
        }

        return `<tr class="${rowClass}">
            <td class="${cellClass}">${concept}</td>
            ${periodCells}
            <td style="font-weight:700; background:rgba(0,0,0,0.02);">${displayAccum}</td>
        </tr>`;
    }).join('');
}

/**
 * 🚀 KPI DASHBOARD: Torre de Control
 */
function renderKPIDashboard(data, selectedIndex) {
    const curr = data[selectedIndex];
    if (!curr) return;

    const kpis = curr.kpis || { ingresos: 0, ebitda: 0, cashflow: 0 };
    const prev = selectedIndex > 0 ? data[selectedIndex - 1] : curr;
    const prevKpis = prev.kpis || kpis;

    // 1. Update Score Cards
    document.getElementById('dash-ingresos').textContent = formatCurrency(kpis.ingresos);
    document.getElementById('dash-ebitda').textContent = formatCurrency(kpis.ebitda);
    
    // Mostramos SALDO FINAL en el Scorecard (Salud)
    const displayCash = kpis.cashEnding || kpis.cashflow;
    document.getElementById('dash-cash').textContent = formatCurrency(displayCash);

    const levValue = curr.balance.ebitdaLTM > 0 ? (curr.balance.deudaTotal / curr.balance.ebitdaLTM).toFixed(2) : "0.00";
    document.getElementById('dash-lev').textContent = levValue + 'x';

    const statusLev = document.getElementById('status-lev');
    if (statusLev) {
        statusLev.textContent = "Estable";
        statusLev.style.color = "var(--text-secondary)";
    }

    // Secondary CEO KPIs
    const utilidad = kpis.utilidad || 0;
    const margenNeto = kpis.margen_neto || 0;
    const margenBruto = kpis.margen_bruto || 0;
    document.getElementById('dash-utilidad').textContent = formatCurrency(utilidad);
    document.getElementById('dash-margen-neto').textContent = formatPercent(margenNeto);
    const margenBrutoEl = document.getElementById('dash-margen-bruto');
    if (margenBrutoEl) margenBrutoEl.textContent = formatPercent(margenBruto);

    // ROE (Utilidad LTM / Patrimonio) - Estimado: multiplicamos utilidad mensual x 12
    const patrimonio = curr.balance.patrimonio > 0 ? curr.balance.patrimonio : 1; 
    const activos = curr.balance.activos > 0 ? curr.balance.activos : 1;
    
    // Si la utilidad y activos son > 0 lo mostramos. Si patrimonio = 0 (anomalía), no mostrar div by zero
    const roe = curr.balance.patrimonio !== 0 ? ((utilidad * 12) / curr.balance.patrimonio) : 0;
    const roa = curr.balance.activos !== 0 ? ((utilidad * 12) / curr.balance.activos) : 0;
    
    document.getElementById('dash-roe').textContent = formatPercent(roe);
    document.getElementById('dash-roa').textContent = formatPercent(roa);

    // CCC = DSO + DIO - DPO
    const dso = curr.cashflowDetail?.dso || 0;
    const dio = curr.cashflowDetail?.dio || 0;
    const dpo = curr.cashflowDetail?.dpo || 0;
    const ccc = dso + dio - dpo;
    
    document.getElementById('dash-ccc').textContent = `${ccc.toFixed(0)} días`;
    
    const prevDso = prev.cashflowDetail?.dso || 0;
    const prevDio = prev.cashflowDetail?.dio || 0;
    const prevDpo = prev.cashflowDetail?.dpo || 0;
    const prevCcc = prevDso + prevDio - prevDpo;
    const cccDiff = prevCcc !== 0 ? ((ccc - prevCcc) / Math.abs(prevCcc)) * 100 : 0;
    if (ccc === 0 && prevCcc === 0) {
        document.getElementById('status-ccc').textContent = "Estable";
        document.getElementById('status-ccc').style.color = "var(--text-secondary)";
    } else if (ccc < prevCcc) {
        document.getElementById('status-ccc').textContent = "▲ Mejorando";
        document.getElementById('status-ccc').style.color = "var(--success)";
    } else if (ccc > prevCcc) {
        document.getElementById('status-ccc').textContent = "▼ Empeorando";
        document.getElementById('status-ccc').style.color = "var(--danger)";
    } else {
        document.getElementById('status-ccc').textContent = "Estable";
        document.getElementById('status-ccc').style.color = "var(--text-secondary)";
    }

    const updateScoreTrend = (id, currVal, prevVal) => {
        const el = document.getElementById(id);
        const diff = currVal - prevVal;
        const pct = prevVal !== 0 ? (diff / Math.abs(prevVal)) * 100 : 0;
        el.style.color = (id === 'trend-cash' || id === 'trend-ebitda' || id === 'trend-ingresos') ? (diff >= 0 ? 'var(--success)' : 'var(--danger)') : 'inherit';
        el.textContent = `${diff >= 0 ? '▲' : '▼'} ${Math.abs(pct).toFixed(1)}% vs mes ant.`;
    };

    updateScoreTrend('trend-ingresos', kpis.ingresos, prevKpis.ingresos);
    updateScoreTrend('trend-ebitda', kpis.ebitda, prevKpis.ebitda);
    updateScoreTrend('trend-utilidad', utilidad, prevKpis.utilidad || 0);
    
    // update simple status for ratios
    const updateRatioStatus = (elId, diff) => {
        const el = document.getElementById(elId);
        if(!el) return;
        el.textContent = diff >= 0 ? "▲ Mejorando" : "▼ Cayendo";
        el.style.color = diff >= 0 ? "var(--success)" : "var(--danger)";
        if (Math.abs(diff) < 0.001) {
            el.textContent = "Estable";
            el.style.color = "var(--text-secondary)";
        }
    };
    
    updateRatioStatus('status-margen-neto', margenNeto - (prevKpis.margen_neto || 0));
    updateRatioStatus('status-margen-bruto', margenBruto - (prevKpis.margen_bruto || 0));

    // ROE, ROA status
    const evaluateStatus = (elId, val) => {
        const el = document.getElementById(elId);
        if(!el) return;
        if(val > 0.15) { el.textContent = "Óptimo"; el.style.color = "var(--success)"; }
        else if(val > 0) { el.textContent = "Adecuado"; el.style.color = "var(--info)"; }
        else if(val === 0) { el.textContent = "Insuficiente Data"; el.style.color = "var(--text-secondary)"; }
        else { el.textContent = "Bajo Cero (Atención)"; el.style.color = "var(--danger)"; }
    };
    evaluateStatus('status-roe', roe);
    evaluateStatus('status-roa', roa);

    const prevDisplayCash = prevKpis.cashEnding || prevKpis.cashflow;
    updateScoreTrend('trend-cash', displayCash, prevDisplayCash);

    // -- Variación Interanual (YoY) --
    let yoyData = null;
    try {
        const prevYearValue = curr.sortDate && !isNaN(new Date(curr.sortDate)) ? new Date(curr.sortDate).getFullYear() - 1 : 2025;
        
        // 1. Try exact month match first
        yoyData = data.find(d => {
            if (d.sortDate && curr.sortDate) {
                const dDate = new Date(d.sortDate);
                const cDate = new Date(curr.sortDate);
                if (!isNaN(dDate) && !isNaN(cDate)) {
                    return dDate.getMonth() === cDate.getMonth() && 
                           dDate.getFullYear() === prevYearValue;
                }
            }
            return false;
        });

        // 2. Fallback: Any data from previous year
        if (!yoyData) {
            yoyData = data.find(d => {
                if (d.sortDate) {
                    const dDate = new Date(d.sortDate);
                    return !isNaN(dDate) && dDate.getFullYear() === prevYearValue;
                }
                const dNorm = normalizeText(d.date || "");
                return dNorm.includes(prevYearValue.toString()) || dNorm.includes(prevYearValue.toString().slice(-2));
            });
        }
    } catch (err) {
        console.warn("Could not find yoyData by date matching.", err);
    }
    
    if (!yoyData) {
        console.warn("Could not find yoyData fallback. Current date:", curr.date, "Available data:", data.map(d=>d.date));
        yoyData = selectedIndex >= 12 ? data[selectedIndex - 12] : null;
    }

    const calcYoY = (currValue, yoyItem, elPrefix) => {
        const valueEl = document.getElementById(`yoy-${elPrefix}`);
        const statusEl = document.getElementById(`yoy-status-${elPrefix}`);
        if (!yoyItem) {
            if (valueEl) valueEl.textContent = "N/A";
            if (statusEl) {
                statusEl.textContent = "Sin datos año ant.";
                statusEl.style.color = "var(--text-secondary)";
            }
            return;
        }
        const prevValue = elPrefix === 'caja' 
            ? (yoyItem.kpis?.cashEnding || yoyItem.kpis?.cashflow || 0)
            : (elPrefix === 'utilidad' ? (yoyItem.kpis?.utilidad || 0) 
            : (yoyItem.kpis?.[elPrefix] || 0));
        
        const diff = currValue - prevValue;
        const pct = prevValue !== 0 ? (diff / Math.abs(prevValue)) * 100 : (currValue !== 0 ? 100 : 0);
        
        if (valueEl) {
            valueEl.textContent = `${pct > 0 ? '+' : ''}${pct.toFixed(1)}%`;
        }
        if (statusEl) {
            if (pct >= 0.01) {
                statusEl.textContent = "▲ Crecimiento";
                statusEl.style.color = "var(--success)";
            } else if (pct <= -0.01) {
                statusEl.textContent = "▼ Decrecimiento";
                statusEl.style.color = "var(--danger)";
            } else {
                statusEl.textContent = "Estable";
                statusEl.style.color = "var(--text-secondary)";
            }
        }
    };
    
    calcYoY(kpis.ingresos, yoyData, 'ingresos');
    calcYoY(kpis.ebitda, yoyData, 'ebitda');
    calcYoY(utilidad, yoyData, 'utilidad');
    calcYoY(displayCash, yoyData, 'caja');
    // --------------------------------

    // 2. Render Sparklines using D3
    const renderSparkline = (containerId, values, color) => {
        const container = document.getElementById(containerId);
        if (!container) return;
        container.innerHTML = '';
        
        const width = container.clientWidth;
        const height = 40;
        const margin = { top: 2, right: 2, bottom: 2, left: 2 };

        const svg = d3.select(`#${containerId}`)
            .append("svg")
            .attr("width", width)
            .attr("height", height);

        const x = d3.scaleLinear()
            .domain([0, Math.max(1, values.length - 1)])
            .range([margin.left, width - margin.right]);

        const y = d3.scaleLinear()
            .domain([d3.min(values) || 0, d3.max(values) || 0])
            .range([height - margin.bottom, margin.top]);

        const line = d3.line()
            .x((d, i) => x(i))
            .y(d => y(d))
            .curve(d3.curveBasis);

        svg.append("path")
            .datum(values)
            .attr("fill", "none")
            .attr("stroke", color)
            .attr("stroke-width", 2)
            .attr("d", line);
    };

    // Filtro: No mostrar 2025 en el Dashboard (Gráficos)
    const rollingData = data.slice(Math.max(0, selectedIndex - 11), selectedIndex + 1).filter(d => !isYear2025(d));
    renderSparkline('spark-ingresos', rollingData.map(d => d.kpis.ingresos), 'var(--success)');
    renderSparkline('spark-ebitda', rollingData.map(d => d.kpis.ebitda), 'var(--primary)');
    renderSparkline('spark-cash', rollingData.map(d => d.kpis.cashflow), 'var(--info)');

    // 3. Main Trend Charts
    renderMarginChart(rollingData);
    renderCashFlowChart(rollingData);

    // 4. Alerts
    renderDashboardAlerts(curr);
}

function renderMarginChart(originalRollingData) {
    const isMobile = window.innerWidth < 768;
    const rollingData = isMobile ? originalRollingData.slice(-3) : originalRollingData;

    const marginContainer = document.getElementById('marginChart');
    if (!marginContainer) return;
    marginContainer.innerHTML = '';

    const width = marginContainer.clientWidth;
    const height = 250;
    const margin = isMobile 
        ? { top: 20, right: 15, bottom: 35, left: 35 } 
        : { top: 20, right: 30, bottom: 40, left: 50 };

    // Tooltip
    let tooltip = d3.select("body").select(".d3-tooltip");
    if (tooltip.empty()) {
        tooltip = d3.select("body").append("div").attr("class", "d3-tooltip");
    }

    const svg = d3.select("#marginChart")
        .append("svg")
        .attr("width", width)
        .attr("height", height)
        .append("g")
        .attr("transform", `translate(${margin.left},${margin.top})`);

    const data = rollingData.map(d => ({
        date: d.date,
        margin: (d.kpis.margen_ebitda || 0) * 100,
        ebitda: d.kpis.ebitda || 0
    }));

    const x = d3.scalePoint()
        .domain(data.map(d => d.date))
        .range([0, width - margin.left - margin.right]);

    const y = d3.scaleLinear()
        .domain([0, d3.max(data, d => d.margin) * 1.2])
        .range([height - margin.top - margin.bottom, 0]);

    svg.append("g")
        .attr("transform", `translate(0,${height - margin.top - margin.bottom})`)
        .call(d3.axisBottom(x).tickSize(0).tickPadding(10))
        .selectAll("text")
        .style("font-size", isMobile ? "8px" : "10px")
        .style("color", "var(--text-secondary)");

    svg.append("g")
        .call(d3.axisLeft(y).ticks(5).tickFormat(d => d + "%"))
        .style("font-size", isMobile ? "8px" : "10px");

    const line = d3.line()
        .x(d => x(d.date))
        .y(d => y(d.margin))
        .curve(d3.curveMonotoneX);

    svg.append("path")
        .datum(data)
        .attr("fill", "none")
        .attr("stroke", "var(--primary)")
        .attr("stroke-width", 3)
        .attr("d", line);

    svg.selectAll(".dot")
        .data(data)
        .enter().append("circle")
        .attr("class", "dot")
        .attr("cx", d => x(d.date))
        .attr("cy", d => y(d.margin))
        .attr("r", 5)
        .attr("fill", "white")
        .attr("stroke", "var(--primary)")
        .attr("stroke-width", 2)
        .style("cursor", "pointer")
        .on("mouseover", function(event, d) {
            d3.select(this).attr("r", 8).attr("fill", "var(--primary)");
            tooltip.style("opacity", 1)
                   .html(`<strong>${d.date}</strong><br/>Margen EBITDA: ${d.margin.toFixed(1)}%<br/>Monto EBITDA: ${formatCurrency(d.ebitda)}`);
        })
        .on("mousemove", function(event) {
            tooltip.style("left", (event.pageX + 10) + "px")
                   .style("top", (event.pageY - 28) + "px");
        })
        .on("mouseout", function() {
            d3.select(this).attr("r", 5).attr("fill", "white");
            tooltip.style("opacity", 0);
        });
}

function renderCashFlowChart(originalRollingData) {
    const isMobile = window.innerWidth < 768;
    const rollingData = isMobile ? originalRollingData.slice(-3) : originalRollingData;

    const cashContainer = document.getElementById('cashFlowChart');
    if (!cashContainer) return;
    cashContainer.innerHTML = '';

    const width = cashContainer.clientWidth;
    const height = 250;
    const margin = isMobile 
        ? { top: 20, right: 15, bottom: 35, left: 35 } 
        : { top: 20, right: 30, bottom: 40, left: 60 };

    // Tooltip
    let tooltip = d3.select("body").select(".d3-tooltip");
    if (tooltip.empty()) {
        tooltip = d3.select("body").append("div").attr("class", "d3-tooltip");
    }

    const svg = d3.select("#cashFlowChart")
        .append("svg")
        .attr("width", width)
        .attr("height", height)
        .append("g")
        .attr("transform", `translate(${margin.left},${margin.top})`);

    const x = d3.scaleBand()
        .domain(rollingData.map(d => d.date))
        .range([0, width - margin.left - margin.right])
        .padding(0.3);

    const y = d3.scaleLinear()
        .domain([d3.min(rollingData, d => d.kpis.cashflow), d3.max(rollingData, d => d.kpis.cashflow) * 1.1])
        .range([height - margin.top - margin.bottom, 0]);

    svg.append("g")
        .attr("transform", `translate(0,${height - margin.top - margin.bottom})`)
        .call(d3.axisBottom(x).tickSize(0).tickPadding(10))
        .selectAll("text")
        .style("font-size", isMobile ? "8px" : "10px");

    svg.append("g")
        .call(d3.axisLeft(y).ticks(5).tickFormat(d => {
            if (d === 0) return "0";
            if (Math.abs(d) >= 1000000) return (d / 1000000).toFixed(1).replace(/\.0$/, '') + "M";
            if (Math.abs(d) >= 1000) return (d / 1000).toFixed(0) + "K";
            return d.toString();
        }))
        .style("font-size", isMobile ? "8px" : "10px");

    svg.selectAll(".bar")
        .data(rollingData)
        .enter().append("rect")
        .attr("class", "bar")
        .attr("x", d => x(d.date))
        .attr("width", x.bandwidth())
        .attr("y", d => y(Math.max(0, d.kpis.cashflow)))
        .attr("height", d => Math.abs(y(d.kpis.cashflow) - y(0)))
        .attr("fill", d => d.kpis.cashflow >= 0 ? "var(--success)" : "var(--danger)")
        .attr("rx", 4)
        .style("cursor", "pointer")
        .on("mouseover", function(event, d) {
            d3.select(this).style("opacity", 0.8);
            tooltip.style("opacity", 1)
                   .html(`<strong>${d.date}</strong><br/>Flujo de Caja: ${formatCurrency(d.kpis.cashflow)}`);
        })
        .on("mousemove", function(event) {
            tooltip.style("left", (event.pageX + 10) + "px")
                   .style("top", (event.pageY - 28) + "px");
        })
        .on("mouseout", function() {
            d3.select(this).style("opacity", 1);
            tooltip.style("opacity", 0);
        });
}

function renderDashboardAlerts(curr) {
    const container = document.getElementById('alertsContainer');
    if (!container) return;
    container.innerHTML = '';

    const alerts = [];
    const kpis = curr.kpis;
    const margin = kpis.margen_ebitda * 100;

    if (margin < 15) {
        alerts.push({ type: 'warning', text: `Margen EBITDA bajo (${margin.toFixed(1)}%). Se recomienda revisar eficiencia operativa.` });
    }
    
    if (curr.integrity && curr.integrity.isBroken) {
        alerts.push({ type: 'danger', text: "Descuadre detectable en la integridad del P&L. Verifique los costos directos." });
    }

    if (curr.balance.activos !== 0 && (curr.balance.activos < curr.balance.pasivos)) {
        alerts.push({ type: 'danger', text: "Patrimonio Negativo detectado. Riesgo de insolvencia técnica." });
    }

    if (alerts.length === 0) {
        container.innerHTML = '<div class="alert-card alert-success">No se detectan anomalías financieras críticas en este periodo.</div>';
    } else {
        container.innerHTML = alerts.map(a => `
            <div class="alert-card alert-${a.type}">
                <i data-lucide="${a.type === 'danger' ? 'alert-octagon' : 'alert-triangle'}"></i>
                <span>${a.text}</span>
            </div>
        `).join('');
    }
    
    if (typeof lucide !== 'undefined') lucide.createIcons();
}

function updateTrend(id, curr, prev, suffix = "") {
    const el = document.getElementById(id);
    const diff = curr - prev;
    const pct = prev !== 0 ? (diff / Math.abs(prev)) * 100 : 0;
    
    if (diff >= 0.01) {
        el.innerHTML = `<span style="color:var(--success)">▲ ${pct.toFixed(1)}%</span> vs mes ant.${suffix}`;
    } else if (diff <= -0.01) {
        el.innerHTML = `<span style="color:var(--danger)">▼ ${Math.abs(pct).toFixed(1)}%</span> vs mes ant.${suffix}`;
    } else {
        el.innerHTML = `Sin cambios vs mes ant.${suffix}`;
    }
}

/**
 * Render Estados Financieros based on wide format
 */
function renderEstadosFinancieros(data, selectedIndex = -1) {
    const headerEl = document.getElementById('header-estados');
    const bodyEl = document.getElementById('body-estados');
    if (!headerEl || !bodyEl || !data || data.length === 0) return;

    const endIdx = selectedIndex >= 0 ? selectedIndex : data.length - 1;
    // Show up to 12 months including the selected one
    const startIdx = Math.max(0, endIdx - 11);
    
    // We do NOT want to show full 12 always if data doesn't have it, but slice will handle that
    const visibleMonths = data.slice(startIdx, endIdx + 1).filter(d => !isYear2025(d));
    const periods = visibleMonths.map(d => d.date);

    // Header
    headerEl.innerHTML = `
        <tr>
            <th style="width: 250px;">Concepto</th>
            ${periods.map(p => `<th style="text-align: right;">${p}</th>`).join('')}
            <th style="text-align: right; background: #f0f9ff; color: #0369a1;">Acumulado YTD</th>
        </tr>
    `;

    let allConcepts = [];
    visibleMonths.forEach(d => {
        if (d.pnl && d.pnl.fullRows) {
            d.pnl.fullRows.forEach(row => {
                if (!allConcepts.includes(row.concept)) allConcepts.push(row.concept);
            });
        }
    });

    const isAccountingRule = (c) => {
        const check = normalizeText(c);
        return check === "activos" || check === "pasivos" || check === "patrimonio" || check === "ventas netas" || 
               check === "utilidad bruta" || check === "ebitda" || check === "ebit" || check === "ingreso antes de impuestos" || 
               check === "beneficio neto" || check === "total activos" || check === "total pasivos" || check === "total patrimonio";
    };
    
    const isCategoryRule = (c) => {
        const check = normalizeText(c);
        return check === "estado de resultados" || check === "estado de situacion" || check === "kpis y drivers" || check === "modulo deuda" || check === "analisis horizontal" || check === "analisis vertical" || check === "analisis margen" || check === "rentabilidad" || check === "variables macro" || check === "balances deuda" || check === "schedule amortizacion" || check === "kpis deuda";
    };

    const isRatio = (c) => {
        const check = normalizeText(c);
        return check.includes("crecimiento") || check.includes("/") || check.includes("margin") || check.includes("margen") || check === "roic" || check === "roe" || check === "roa" || check === "dscr" || check.includes("tasa");
    };

    const isDecimal = (c) => {
        const check = normalizeText(c);
        return check.includes("fx ") || check.includes("tasa") || check === "leverage";
    };

    let tbBody = '';
    allConcepts.forEach(concept => {
        if (normalizeText(concept) === "x" || normalizeText(concept) === "año" || normalizeText(concept) === "mes" || normalizeText(concept) === "columna" || normalizeText(concept) === "(dop)") return; 
        
        let headerColor = isCategoryRule(concept) ? 'background: #e0f2fe; color: #0369a1; font-weight: 800; font-size: 1.1em;' : (isAccountingRule(concept) ? 'background: #f8fafc; font-weight: 700;' : '');
        let rowHtml = `<td style="font-weight: ${isAccountingRule(concept) || isCategoryRule(concept) ? '700' : '500'}; color: ${isAccountingRule(concept) || isCategoryRule(concept) ? 'var(--sidebar)' : 'var(--text-primary)'};">${concept}</td>`;
        let total = 0;
        let isTotalizable = true;
        let anyVal = false;

        periods.forEach(p => {
            const periodData = visibleMonths.find(d => d.date === p);
            let val = 0;
            if (periodData && periodData.pnl && periodData.pnl.fullRows) {
                const foundRow = periodData.pnl.fullRows.find(r => r.concept === concept);
                if (foundRow && foundRow.values[p] !== undefined) {
                    val = foundRow.values[p];
                    if (val !== 0) anyVal = true;
                }
            }
            if (isCategoryRule(concept)) {
                rowHtml += `<td></td>`;
            } else if (isRatio(concept)) {
                isTotalizable = false;
                rowHtml += `<td style="text-align: right; font-family: 'JetBrains Mono', monospace;">${val === 0 ? '-' : formatPercent(val)}</td>`;
            } else if (isDecimal(concept)) {
                isTotalizable = false;
                rowHtml += `<td style="text-align: right; font-family: 'JetBrains Mono', monospace;">${val === 0 ? '-' : val.toLocaleString('en-US', {minimumFractionDigits: 2, maximumFractionDigits: 2})}</td>`;
            } else {
                rowHtml += `<td style="text-align: right; font-family: 'JetBrains Mono', monospace;">${val === 0 ? '-' : formatRawCurrency(val)}</td>`;
                total += val;
            }
        });

        if (isCategoryRule(concept)) {
            rowHtml += `<td style="background: #e0f2fe;"></td>`;
        } else if (isTotalizable) {
            rowHtml += `<td style="text-align: right; font-family: 'JetBrains Mono', monospace; font-weight: 700; background: #f0f9ff; color: #0369a1;">${total === 0 ? '-' : formatRawCurrency(total)}</td>`;
        } else {
            rowHtml += `<td style="text-align: right; font-family: 'JetBrains Mono', monospace; font-weight: 700; background: #f0f9ff; color: #0369a1;">-</td>`;
        }

        // Only show if there was any value or if it's a category
        if (anyVal || isCategoryRule(concept)) {
            tbBody += `<tr style="${headerColor}">${rowHtml}</tr>`;
        }
    });

    bodyEl.innerHTML = tbBody;
}
