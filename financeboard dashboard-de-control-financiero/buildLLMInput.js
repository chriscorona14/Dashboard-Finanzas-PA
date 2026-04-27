/**
 * 🧠 Generador de Input para el LLM
 */
export function buildLLMInput({ pnlData, balanceData, source }) {

    function buildPnLRows(pnlData) {
        if (!pnlData) return [];

        return [
            {
                label: "Ventas Netas",
                values: pnlData.ventas || []
            },
            {
                label: "EBITDA",
                values: pnlData.ebitda || []
            }
        ];
    }

    function buildBalanceRows(balanceData) {
        if (!balanceData) return [];

        return [
            {
                label: "Activos",
                values: [balanceData.activos || 0]
            },
            {
                label: "Pasivos",
                values: [balanceData.pasivos || 0]
            },
            {
                label: "Patrimonio",
                values: [balanceData.patrimonio || 0]
            }
        ];
    }

    return {
        instruction: "Process this financial data following system rules. Detect model type automatically and prioritize financial statements if present.",

        metadata: {
            source: source || "excel_upload",
            expected_output: "financial_dashboard",
            strict_mode: true
        },

        data: {
            financial_statements: {
                pnl: {
                    sheet_name: "P&L",
                    rows: buildPnLRows(pnlData)
                },
                balance: {
                    sheet_name: "Balance",
                    rows: buildBalanceRows(balanceData)
                }
            },
            fallback: {
                tb: [],
                setup: []
            }
        },

        rules_override: {
            use_financial_statements_first: true,
            prevent_recalculation: true,
            ignore_totals: true
        },

        output_constraints: {
            format: "json_only",
            no_explanations: true
        }
    };
}
