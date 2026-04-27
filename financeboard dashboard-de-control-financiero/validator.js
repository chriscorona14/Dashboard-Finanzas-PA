/**
 * 🔍 Validador de Input para el LLM
 */
export function validateLLMInput(input) {
    const errors = [];

    const pnlRows = input?.data?.financial_statements?.pnl?.rows || [];
    const balanceRows = input?.data?.financial_statements?.balance?.rows || [];

    // 🔴 Validar P&L
    if (!pnlRows.length) {
        errors.push("P&L vacío");
    }

    const ventas = pnlRows.find(r => r.label.toLowerCase().includes("ventas"));
    const ebitda = pnlRows.find(r => r.label.toLowerCase().includes("ebitda"));

    if (!ventas || !ventas.values.length) {
        errors.push("Ventas no detectadas");
    }

    if (!ebitda || !ebitda.values.length) {
        errors.push("EBITDA no detectado");
    }

    // 🔴 Validar Balance
    if (balanceRows.length) {
        const activos = balanceRows.find(r => r.label.toLowerCase().includes("activos"));
        const pasivos = balanceRows.find(r => r.label.toLowerCase().includes("pasivos"));
        const patrimonio = balanceRows.find(r => r.label.toLowerCase().includes("patrimonio"));

        if (!activos || !pasivos || !patrimonio) {
            errors.push("Balance incompleto");
        }

        const a = activos?.values?.[0] || 0;
        const p = pasivos?.values?.[0] || 0;
        const pat = patrimonio?.values?.[0] || 0;

        // Tolerancia para el cuadre
        if (Math.abs(a - (p + pat)) > 100) {
            errors.push("Balance no cuadra");
        }
    }

    return {
        isValid: errors.length === 0,
        errors
    };
}
