import ExcelJS from "exceljs";

export interface Rule002Metadata {
    fromMonth: number;
    toMonth: number;
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

export class Rule002Exporter {
    async export(results: any[]) {
        const workbook = new ExcelJS.Workbook();
        workbook.creator = "HM Kardex Audit";
        workbook.created = new Date();
        const worksheet = workbook.addWorksheet("RULE_002");
        worksheet.mergeCells("A1:P1");
        worksheet.getCell("A1").value = "RULE_002 - Validación de Continuidad Mensual";
        worksheet.getCell("A1").font = { bold: true, size: 16};
        worksheet.addRow([]);
        /*worksheet.columns = [
            { header: "Código", key: "code", width: 18 },
            { header: "Producto", key: "product", width: 40 },
            { header: "Mes Final", key: "fromMonth", width: 12 },
            { header: "Mes Inicial", key: "toMonth", width: 12 },
            { header: "Cantidad Final", key: "finalQuantity", width: 18 },
            { header: "Cantidad Inicial", key: "initialQuantity", width: 18 },
            { header: "Diferencia Cantidad", key: "quantityDifference", width: 20 },
            { header: "Costo Unitario Final", key: "finalUnitCost", width: 20 },
            { header: "Costo Unitario Inicial", key: "initialUnitCost", width: 20 },
            { header: "Diferencia Costo Unitario", key: "unitCostDifference", width: 22 },
            { header: "Costo Total Final", key: "finalTotalCost", width: 20 },
            { header: "Costo Total Inicial", key: "initialTotalCost", width: 20 },
            { header: "Diferencia Costo Total", key: "totalCostDifference", width: 22 },
            { header: "Campos con Diferencia", key: "differences", width: 30 },
            { header: "Descripción", key: "description", width: 50 },
            { header: "Recomendación", key: "recommendation", width: 50 }
        ];*/
        worksheet.addRow([
            "Código",
            "Producto",
            "Mes Final",
            "Mes Inicial",
            "Cantidad Final",
            "Cantidad Inicial",
            "Diferencia Cantidad",
            "Costo Unitario Final",
            "Costo Unitario Inicial",
            "Diferencia Costo Unitario",
            "Costo Total Final",
            "Costo Total Inicial",
            "Diferencia Costo Total",
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
        worksheet.getColumn(9).width = 20;
        worksheet.getColumn(10).width = 22;
        worksheet.getColumn(11).width = 20;
        worksheet.getColumn(12).width = 20;
        worksheet.getColumn(13).width = 22;
        worksheet.getColumn(14).width = 30;
        worksheet.getColumn(15).width = 50;
        worksheet.getColumn(16).width = 50;

        for (const result of results) {
            const metadata = result.metadata as Rule002Metadata;
            worksheet.addRow([
                result.product_code,
                result.product_name,
                metadata.fromMonth,
                metadata.toMonth,
                metadata.finalBalance.quantity,
                metadata.initialBalance.quantity,
                metadata.finalBalance.quantity -
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
            to: "P3"
        };

        return workbook;

    }

}