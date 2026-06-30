import ExcelJS from "exceljs";

export interface Rule001Metadata {
    inventoryCode: string;
    normalizedCode: string;
    inventoryStock: number;
    kardexStock: number;
    inventoryUnitCost: number;
    kardexUnitCost: number;
    inventoryTotalCost: number;
    kardexTotalCost: number;
    kardexMovements: number;
}

export class Rule001Exporter {

    async export(results: any[]) {

        const workbook = new ExcelJS.Workbook();

        workbook.creator = "HM Kardex Audit";
        workbook.created = new Date();

        const worksheet = workbook.addWorksheet("RULE_001");

        worksheet.mergeCells("A1:N1");

        worksheet.getCell("A1").value =
            "RULE_001 - Validación Inventario Final vs Kardex";

        worksheet.getCell("A1").font = {
            bold: true,
            size: 16
        };

        // Espacio entre título y tabla
        worksheet.addRow([]);

        // Cabeceras
        worksheet.addRow([
            "Código",
            "Producto",
            "Stock Inventario",
            "Stock Kardex",
            "Diferencia Stock",
            "Costo Unitario Inventario",
            "Costo Unitario Kardex",
            "Diferencia Costo Unitario",
            "Costo Total Inventario",
            "Costo Total Kardex",
            "Diferencia Costo Total",
            "Movimientos Kardex",
            "Descripción",
            "Recomendación"
        ]);

        // Anchos
        worksheet.getColumn(1).width = 18;
        worksheet.getColumn(2).width = 40;
        worksheet.getColumn(3).width = 18;
        worksheet.getColumn(4).width = 18;
        worksheet.getColumn(5).width = 18;
        worksheet.getColumn(6).width = 22;
        worksheet.getColumn(7).width = 22;
        worksheet.getColumn(8).width = 22;
        worksheet.getColumn(9).width = 22;
        worksheet.getColumn(10).width = 22;
        worksheet.getColumn(11).width = 22;
        worksheet.getColumn(12).width = 18;
        worksheet.getColumn(13).width = 60;
        worksheet.getColumn(14).width = 60;

        // Datos
        for (const result of results) {

            const metadata = result.metadata as Rule001Metadata;

            worksheet.addRow([
                result.product_code,
                result.product_name,

                metadata.inventoryStock,
                metadata.kardexStock,
                metadata.inventoryStock - metadata.kardexStock,

                metadata.inventoryUnitCost,
                metadata.kardexUnitCost,
                metadata.inventoryUnitCost - metadata.kardexUnitCost,

                metadata.inventoryTotalCost,
                metadata.kardexTotalCost,
                metadata.inventoryTotalCost - metadata.kardexTotalCost,

                metadata.kardexMovements,

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
            to: "N3"
        };

        return workbook;

    }

}