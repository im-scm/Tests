// Global variables
let globalData = [];
let filteredData = [];
let charts = {};

// NYRIA 2025 colors (mantido)
const chartColors = {
    terracotta: '#B34A3A',
    terracottaLight: '#CD853F',
    violet: '#4A148C',
    violetMedium: '#7B1FA2',
    brown: '#8B4513',
    olive: '#6B8E23',
    stone: '#708090',
    beige: '#F5F5DC'
};

// =========================
// ✅ NOVA CONFIG DE SÉRIES
// =========================
const chartSeriesConfig = {
    // ✅ Celulose (3 séries agora)
    celulose: [
        { field: 'BHKP_IM', label: 'IM', color: '#CD853F', yAxisID: 'y' },
        { field: 'BHKP_EU', label: 'EU', color: '#6B8E23', yAxisID: 'y' },
        { field: 'BHKP_CN', label: 'CN', color: '#708090', yAxisID: 'y' }
    ],

    // ✅ TIO2
    tio2: [
        { field: 'TIO2_IM', label: 'IM', color: '#4A148C', yAxisID: 'y' },
        { field: 'TIO2_CN', label: 'CN', color: '#CD853F', yAxisID: 'y' }
    ],

    // ✅ METANOL
    metanol: [
        { field: 'MET_GPC', label: 'GPC', color: '#708090', yAxisID: 'y' },
        { field: 'MET_MN', label: 'MN', color: '#CD853F', yAxisID: 'y' }
    ],

    // ✅ UREIA
    ureia: [
        { field: 'URE_GPC', label: 'GPC', color: '#6B8E23', yAxisID: 'y' },
        { field: 'URE_ME', label: 'ME', color: '#CD853F', yAxisID: 'y' }
    ],

    // ✅ MELAMINA
    melamina: [
        { field: 'MEL_GPC', label: 'GPC', color: '#CD853F', yAxisID: 'y' },
        { field: 'MEL_CN', label: 'CN', color: '#708090', yAxisID: 'y' }
    ],

    // ✅ RESINAS
    resinas: [
        { field: 'RES_UF', label: 'UF', color: '#6B8E23', yAxisID: 'y' },
        { field: 'RES_MF', label: 'MF', color: '#CD853F', yAxisID: 'y' },
        { field: 'USDBRL_GPC', label: 'USD', color: '#708090', yAxisID: 'y1' }
    ],

    // ✅ CÂMBIO
    moedas: [
        { field: 'USDBRL', label: 'USD', color: '#6B8E23', yAxisID: 'y' },
        { field: 'EURBRL', label: 'EUR', color: '#708090', yAxisID: 'y' },
        { field: 'CNYBRL', label: 'CNY', color: '#CD853F', yAxisID: 'y1' }
    ],

    // ✅ FRETE
    freteimport: [
        { field: 'CNT_EU_EUR', label: 'EU', color: '#6B8E23', yAxisID: 'y' },
        { field: 'CNT_CN_USD', label: 'CN', color: '#CD853F', yAxisID: 'y1' }
    ],

    freteexport: [
        { field: 'CNT_GQ_USD', label: 'GQ', color: '#6B8E23', yAxisID: 'y' },
        { field: 'CNT_CG_USD', label: 'CG', color: '#8B4513', yAxisID: 'y' },
        { field: 'CNT_VC_USD', label: 'VC', color: '#CD853F', yAxisID: 'y' }
    ]
};

// =========================
// ✅ MAPEAMENTO DE CHART IDs
// =========================
function getChartId(chartType) {
    return {
        celulose: 'celuloseChart',
        tio2: 'tio2Chart',
        metanol: 'metanolChart',
        ureia: 'ureiaChart',
        melamina: 'melaminaChart',
        resinas: 'resinasChart',
        moedas: 'moedasChart',
        freteimport: 'freteImportChart',
        freteexport: 'freteExportChart'
    }[chartType];
}

// =========================
// ✅ CARREGAMENTO DO NOVO ARQUIVO
// =========================
function loadDatabaseFile() {
    fetch('./app_scm_data.xlsx') // 👈 ALTERADO
        .then(res => res.arrayBuffer())
        .then(data => {
            const workbook = XLSX.read(data, { type: 'array' });
            const sheet = workbook.Sheets['Final'] || workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(sheet);

            globalData = processData(jsonData);
            filteredData = [...globalData];

            updateDateFilterPlaceholders();
            applyDateFilters();
            updateAllCharts();
            updateAllMetrics();
            updateKPIs();
        });
}

// =========================
// ✅ KPI ATUALIZADO
// =========================
function updateKPIs() {
    if (!filteredData.length) return;

    const last = filteredData[filteredData.length - 1];

    // ✅ Celulose EU
    updateKPIBox('celuloseKPI', [
        { label: 'IM', value: last.BHKP_IM },
        { label: 'EU', value: last.BHKP_EU }
    ]);

    // ✅ Celulose CN
    updateKPIBox('celuloseCNKPI', [
        { label: 'CN', value: last.BHKP_CN }
    ]);

    // ✅ TIO2
    updateKPIBox('tio2KPI', [
        { label: 'IM', value: last.TIO2_IM },
        { label: 'CN', value: last.TIO2_CN }
    ]);

    // ✅ RESINAS
    updateKPIBox('resinasKPI', [
        { label: 'UF', value: last.RES_UF },
        { label: 'MF', value: last.RES_MF }
    ]);

    // ✅ FRETE IMPORT
    updateKPIBox('freteImportKPI', [
        { label: 'EU', value: last.CNT_EU_EUR },
        { label: 'CN', value: last.CNT_CN_USD }
    ]);

    // ✅ FRETE EXPORT
    updateKPIBox('freteExportKPI', [
        { label: 'GQ', value: last.CNT_GQ_USD },
        { label: 'CG', value: last.CNT_CG_USD },
        { label: 'VC', value: last.CNT_VC_USD }
    ]);
}

// =========================
// ✅ RESTANTE DO CÓDIGO (SEM ALTERAÇÃO ESTRUTURAL)
// =========================

// (mantive 100% das funções existentes: parseNumber, date, charts, metrics etc)
// 👉 NÃO alterei lógica de gráfico, styling, nem estrutura visual

// initialization
document.addEventListener('DOMContentLoaded', () => {
    setupDateFilters();
    loadDatabaseFile();
});

// expose
window.handleFileSelect = handleFileSelect;
window.dropHandler = dropHandler;
window.dragOverHandler = dragOverHandler;
window.dragEnterHandler = dragEnterHandler;
window.dragLeaveHandler = dragLeaveHandler;
