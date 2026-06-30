import ExcelJS from "exceljs";

export interface Rule005Metadata {
    source: string;

    // Inventario
    stock?: number;
    unitCost?: number;
    totalCost?: number;

    // Kardex
    date?: string;
    month?: number;
    document?: string;
    operation?: string;

    balanceQuantity?: number;
    balanceUnitCost?: number;
    balanceTotalCost?: number;

    entryQuantity?: number;
    entryUnitCost?: number;
    entryTotalCost?: number;

    exitQuantity?: number;
    exitUnitCost?: number;
    exitTotalCost?: number;

    negatives?: string[];
}

export class Rule005Exporter {

    async export(results: any[]) {

        const workbook = new ExcelJS.Workbook();

        workbook.creator = "HM Kardex Audit";
        workbook.created = new Date();

        const worksheet = workbook.addWorksheet("RULE_005");

        worksheet.mergeCells("A1:V1");

        worksheet.getCell("A1").value =
            "RULE_005 - Validación de Saldos Negativos";

        worksheet.getCell("A1").font = {
            bold: true,
            size: 16
        };

        // Espacio
        worksheet.addRow([]);

        // Cabeceras
        worksheet.addRow([
            "Origen",
            "Código",
            "Producto",
            "Mes",
            "Fecha",
            "Documento",
            "Operación",
            "Stock",
            "Costo Unitario",
            "Costo Total",
            "Saldo Cantidad",
            "Saldo Costo Unitario",
            "Saldo Costo Total",
            "Ingreso Cantidad",
            "Ingreso Costo Unitario",
            "Ingreso Costo Total",
            "Salida Cantidad",
            "Salida Costo Unitario",
            "Salida Costo Total",
            "Campos Negativos",
            "Descripción",
            "Recomendación"
        ]);

        // Anchos
        worksheet.getColumn(1).width = 15;
        worksheet.getColumn(2).width = 18;
        worksheet.getColumn(3).width = 40;
        worksheet.getColumn(4).width = 10;
        worksheet.getColumn(5).width = 15;
        worksheet.getColumn(6).width = 20;
        worksheet.getColumn(7).width = 20;
        worksheet.getColumn(8).width = 15;
        worksheet.getColumn(9).width = 18;
        worksheet.getColumn(10).width = 18;
        worksheet.getColumn(11).width = 18;
        worksheet.getColumn(12).width = 20;
        worksheet.getColumn(13).width = 20;
        worksheet.getColumn(14).width = 18;
        worksheet.getColumn(15).width = 20;
        worksheet.getColumn(16).width = 20;
        worksheet.getColumn(17).width = 18;
        worksheet.getColumn(18).width = 20;
        worksheet.getColumn(19).width = 20;
        worksheet.getColumn(20).width = 35;
        worksheet.getColumn(21).width = 55;
        worksheet.getColumn(22).width = 55;

        // Datos
        for (const result of results) {

            const metadata = result.metadata as Rule005Metadata;

            worksheet.addRow([

                metadata.source,

                result.product_code,

                result.product_name,

                metadata.month,

                metadata.date,

                metadata.document,

                metadata.operation,

                metadata.stock,

                metadata.unitCost,

                metadata.totalCost,

                metadata.balanceQuantity,

                metadata.balanceUnitCost,

                metadata.balanceTotalCost,

                metadata.entryQuantity,

                metadata.entryUnitCost,

                metadata.entryTotalCost,

                metadata.exitQuantity,

                metadata.exitUnitCost,

                metadata.exitTotalCost,

                metadata.negatives?.join(", "),

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
            to: "V3"
        };

        return workbook;

    }

}