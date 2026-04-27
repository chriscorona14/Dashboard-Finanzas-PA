import * as XLSX from 'xlsx';
import { GoogleGenAI } from "@google/genai";
import * as d3 from 'd3';
import { financialEngine, formatCurrency, formatRawCurrency, formatPercent, normalizeText } from "./financialEngine.js";
import { buildLLMInput } from "./buildLLMInput.js";
import { validateLLMInput } from "./validator.js";

// --- ESTADO GLOBAL ---
let globalFinancialData = [];
const loader = document.getElementById('loader');
const monthSelector = document.getElementById('monthSelector');

// --- INICIALIZACIÓN DE GEMINI ---
// Importante: Usamos import.meta.env para compatibilidad con Vite/Vercel
const ai = new GoogleGenAI({ 
    apiKey: import.meta.env.VITE_GEMINI_API_KEY 
});

// --- CONFIGURACIÓN MICROSOFT MSAL ---
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

// --- FUNCIONES DE CONEXIÓN ---

async function connectM365() {
    if (!msalInstance) {
        alert("MSAL no inicializado.");
        return;
    }

    try {
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
            alert("Hay una autenticación en progreso. Por favor, abra la app en una nueva pestaña.");
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
        statusEl.innerHTML = "⏳ Sincronizando modelo remoto...";
    }
    if (loader) loader.style.display = 'flex';
    
    const loginBtn = document.getElementById('loginM365Btn');
    if (loginBtn) loginBtn.style.display = 'none';

    try {
        let arrayBuffer;
        
        if (token) {
            const encodedUrl = btoa(SHARPOINT_FILE_URL).replace(/=/g, '').replace(/\//g, '_').replace(/\+/g, '-');
            const graphUrl = `https://graph.microsoft.com/v1.0/shares/u!${encodedUrl}/driveItem/content`;
            
            const req = await fetch(graphUrl, {
                headers: { "Authorization": `Bearer ${token}` }
            });
            if (!req.ok) throw new Error(`O365 Graph Error: ${req.status}`);
            arrayBuffer = await req.arrayBuffer();
        } else {
            const response = await fetch("/api/downloadSync");
            if (!response.ok) throw new Error("Error en Proxy");
            arrayBuffer = await response.arrayBuffer();
        }
        
        const workbook = XLSX.read(new Uint8Array(arrayBuffer), { type: 'array', cellDates: true });
        const engineResult = financialEngine(workbook);
        
        if (engineResult.error || !engineResult.data) throw new Error(engineResult.error);
        
        if (statusEl) statusEl.innerHTML = "✅ Sincronizado con O365";
        if (sidebarSyncDot) sidebarSyncDot.style.backgroundColor = 'var(--success)';
        if (sidebarSyncText) {
            sidebarSyncText.innerText = 'Sincronizado';
            sidebarSyncText.style.color = 'var(--success)';
        }
        
        globalFinancialData = engineResult.data;
        renderDashboard(globalFinancialData);
        if (loader) loader.style.display = 'none';
        
    } catch (error) {
        console.error("Auto-sync failed:", error);
        if (statusEl) {
            statusEl.style.background = '#fee2e2';
            statusEl.style.color = '#991b1b';
            statusEl.innerHTML = "⚠️ Conexión fallida. Use carga manual.";
        }
        if (sidebarSyncDot) sidebarSyncDot.style.backgroundColor = 'var(--danger)';
        if (sidebarSyncText) {
             sidebarSyncText.innerText = 'Desconectado';
             sidebarSyncText.style.color = 'var(--danger)';
        }
        if (loader) loader.style.display = 'none';
        if (loginBtn) loginBtn.style.display = 'flex';
    }
}

// --- LOGICA DE NAVEGACIÓN ---

/**
 * Función central para cambiar entre hojas del dashboard.
 * Se expone a window para que funcione desde el HTML y la consola.
 */
function showSection(sectionId) {
    const titles = {
        'view-kpi': "Torre de Control: Indicadores Clave",
        'view-resumen': "Dashboard de Gestión Corporativa (RD$)",
        'view-pnl': "Estado de Resultados Detallado (RD$)",
        'view-balance': "Balance General Consolidado (RD$)",
        'view-cashflow': "Estado de Flujo de Efectivo (RD$)",
        'view-config': "Configuración y Auditoría",
        'view-estados': "Estados Financieros",
        'view-glosario': "Glosario Financiero"
    };

    // Actualizar secciones activas
    document.querySelectorAll('.view-container').forEach(v => v.classList.remove('active'));
    const target = document.getElementById(sectionId);
    if (target) target.classList.add('active');

    // Actualizar título
    const titleLabel = document.getElementById('titleLabel');
    if (titleLabel && titles[sectionId]) titleLabel.textContent = titles[sectionId];

    // Control de visibilidad de controles superiores
    const periodContainer = document.getElementById('periodContainer');
    const searchWrapper = document.getElementById('searchContainerWrapper');
    
    const isConfigOrGlosario = sectionId === 'view-config' || sectionId === 'view-glosario';
    if (periodContainer) periodContainer.style.display = isConfigOrGlosario ? 'none' : 'flex';
    
    if (monthSelector) {
        monthSelector.style.display = (isConfigOrGlosario || !globalFinancialData.length) ? 'none' : 'block';
    }

    // Forzar redibujado de gráficos para ajustar dimensiones
    if (globalFinancialData.length > 0) {
        const idx = parseInt(monthSelector.value);
        if (!isNaN(idx)) updateUI(globalFinancialData, idx);
    }
    
    // Cerrar sidebar en móviles tras click
    const sidebar = document.querySelector('.sidebar');
    if (sidebar && window.innerWidth <= 1024) {
        sidebar.classList.remove('open');
    }
}

// --- INICIALIZACIÓN DEL DOM ---

document.addEventListener('DOMContentLoaded', () => {
    fetchMasterData();
    
    // Listeners para botones de exportación y carga
    document.getElementById('loginM365Btn')?.addEventListener('click', connectM365);
    document.getElementById('fileInput')?.addEventListener('change', handleFileUpload);
    
    // Listener para el selector de meses
    monthSelector?.addEventListener('change', (e) => {
        const index = parseInt(e.target.value);
        if (!isNaN(index)) updateUI(globalFinancialData, index);
    });

    // Delegación de eventos para la barra lateral (Sidebar)
    const menuItems = ['menu-kpi', 'menu-resumen', 'menu-pnl', 'menu-balance', 'menu-cashflow', 'menu-estados', 'menu-config', 'menu-glosario'];
    menuItems.forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.addEventListener('click', (e) => {
                e.preventDefault();
                menuItems.forEach(mId => document.getElementById(mId)?.classList.remove('active'));
                el.classList.add('active');
                showSection(id.replace('menu-', 'view-'));
            });
        }
    });

    if (typeof lucide !== 'undefined') lucide.createIcons();
});

// --- EXPOSICIÓN GLOBAL PARA VITE/VERCEL ---
window.showSection = showSection;
window.connectM365 = connectM365;
window.handleFileUpload = handleFileUpload;

// --- (RESTO DE TUS FUNCIONES DE RENDERIZADO: renderDashboard, updateUI, etc.) ---
// Asegúrate de copiar todas las funciones que siguen (renderKPIDashboard, renderDetailedPnL, etc.) debajo de estas líneas.
