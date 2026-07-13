import ExcelJS from "exceljs";

export interface Rule013Metadata {
    month: number;
    normalizedCode: string;

    totals: {
        entry: {
            quantity: number;
            totalCost: number;
        };
        exit: {
            quantity: number;
            totalCost: number;
        };
    };

    initialBalance: {
        quantity: number;
        totalCost: number;
    };

    expectedFinalBalance: {
        quantity: number;
        totalCost: number;
    };

    actualFinalBalance: {
        quantity: number;
        totalCost: number;
    };

    movementCount: number;
    differences: string[];
}

export class Rule013Exporter {

    async export(results: any[]) {

        const workbook = new ExcelJS.Workbook();

        workbook.creator = "HM Kardex Audit";
        workbook.created = new Date();

        const worksheet = workbook.addWorksheet("RULE_013");

        worksheet.mergeCells("A1:P1");

        worksheet.getCell("A1").value =
            "RULE_013 - Validación de Sumatorias Mensuales";

        worksheet.getCell("A1").font = {
            bold: true,
            size: 16
        };

        // Espacio
        worksheet.addRow([]);

        // Cabeceras
        worksheet.addRow([
            "Código",
            "Producto",
            "Mes",

            "Saldo Inicial - Unidades",
            "Saldo Inicial - Costo Valorizado",

            "Total Entradas - Unidades",
            "Total Entradas - Costo Valorizado",

            "Total Salidas - Unidades",
            "Total Salidas - Costo Valorizado",

            "Saldo Final Esperado - Unidades",
            "Saldo Final Esperado - Costo Valorizado",

            "Saldo Final Real - Unidades",
            "Saldo Final Real - Costo Valorizado",

            "Diferencias",
            "Descripción",
            "Recomendación"
        ]);

        // Anchos
        worksheet.getColumn(1).width = 18;
        worksheet.getColumn(2).width = 40;
        worksheet.getColumn(3).width = 10;

        worksheet.getColumn(4).width = 25;
        worksheet.getColumn(5).width = 30;

        worksheet.getColumn(6).width = 25;
        worksheet.getColumn(7).width = 30;

        worksheet.getColumn(8).width = 25;
        worksheet.getColumn(9).width = 30;

        worksheet.getColumn(10).width = 30;
        worksheet.getColumn(11).width = 35;

        worksheet.getColumn(12).width = 28;
        worksheet.getColumn(13).width = 35;

        worksheet.getColumn(14).width = 25;
        worksheet.getColumn(15).width = 55;
        worksheet.getColumn(16).width = 55;

        // Datos
        for (const result of results) {

            const metadata = result.metadata as Rule013Metadata;

            worksheet.addRow([
                result.product_code,
                result.product_name,
                metadata.month,

                metadata.initialBalance.quantity,
                metadata.initialBalance.totalCost,

                metadata.totals.entry.quantity,
                metadata.totals.entry.totalCost,

                metadata.totals.exit.quantity,
                metadata.totals.exit.totalCost,

                metadata.expectedFinalBalance.quantity,
                metadata.expectedFinalBalance.totalCost,

                metadata.actualFinalBalance.quantity,
                metadata.actualFinalBalance.totalCost,

                metadata.differences.join(", "),
                result.description,
                result.recommendation
            ]);
        }

        worksheet.views = [
            {
                state: "frozen",
                ySplit: 3
            }
        ];

        worksheet.autoFilter = {
            from: "A3",
            to: "P3"
        };

        return workbook;
    }
}