import ExcelJS from "exceljs";

export interface Rule003Metadata {
    fromIndex: number;
    toIndex: number;
    finalBalance: {
        quantity: number;
        unitCost: number;
        totalCost: number;
    };
    initialBalance: {
        quantity: number;
        unitCost: number;
        totalCost: number;
    };
    differences: string[];
}

export class Rule003Exporter {

    async export(results: any[]) {

        const workbook = new ExcelJS.Workbook();

        workbook.creator = "HM Kardex Audit";
        workbook.created = new Date();

        const worksheet = workbook.addWorksheet("RULE_003");

        worksheet.mergeCells("A1:O1");

        worksheet.getCell("A1").value =
            "RULE_003 - Validación de Continuidad de Costos";

        worksheet.getCell("A1").font = {
            bold: true,
            size: 16
        };

        worksheet.addRow([]);

        worksheet.addRow([
            "Código",
            "Producto",
            "Mes Cierre",
            "Mes Inicio",
            "Cantidad Final",
            "Cantidad Inicial",
            "Costo Unitario Final",
            "Costo Unitario Inicial",
            "Dif. Costo Unitario",
            "Costo Total Final",
            "Costo Total Inicial",
            "Dif. Costo Total",
            "Campos con Diferencia",
            "Descripción",
            "Recomendación"
        ]);

        worksheet.getColumn(1).width = 18;
        worksheet.getColumn(2).width = 40;
        worksheet.getColumn(3).width = 12;
        worksheet.getColumn(4).width = 12;
        worksheet.getColumn(5).width = 18;
        worksheet.getColumn(6).width = 18;
        worksheet.getColumn(7).width = 20;
        worksheet.getColumn(8).width = 20;
        worksheet.getColumn(9).width = 22;
        worksheet.getColumn(10).width = 20;
        worksheet.getColumn(11).width = 20;
        worksheet.getColumn(12).width = 22;
        worksheet.getColumn(13).width = 35;
        worksheet.getColumn(14).width = 50;
        worksheet.getColumn(15).width = 50;

        for (const result of results) {

            const metadata = result.metadata as Rule003Metadata;

            worksheet.addRow([
                result.product_code,
                result.product_name,
                metadata.fromIndex,
                metadata.toIndex,
                metadata.finalBalance.quantity,
                metadata.initialBalance.quantity,
                metadata.finalBalance.unitCost,
                metadata.initialBalance.unitCost,
                metadata.finalBalance.unitCost -
                    metadata.initialBalance.unitCost,
                metadata.finalBalance.totalCost,
                metadata.initialBalance.totalCost,
                metadata.finalBalance.totalCost -
                    metadata.initialBalance.totalCost,
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
            to: "O3"
        };

        return workbook;

    }

}