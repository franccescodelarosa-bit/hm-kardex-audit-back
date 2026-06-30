import ExcelJS from "exceljs";

export interface Rule009Metadata {
    date: string;
    month: number;
    document: string;
    operation: string;

    entryQuantity: number;
    entryUnitCost: number;
    entryTotalCost: number;

    balanceQuantity: number;
    balanceUnitCost: number;
    balanceTotalCost: number;
}

export class Rule009Exporter {

    async export(results: any[]) {

        const workbook = new ExcelJS.Workbook();

        workbook.creator = "HM Kardex Audit";
        workbook.created = new Date();

        const worksheet = workbook.addWorksheet("RULE_009");

        worksheet.mergeCells("A1:N1");

        worksheet.getCell("A1").value =
            "RULE_009 - Ingresos por Ajuste de Inventario";

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
            "Fecha",
            "Documento",
            "Operación",
            "Cantidad Ingreso",
            "Costo Unitario Ingreso",
            "Costo Total Ingreso",
            "Saldo Cantidad",
            "Saldo Costo Unitario",
            "Saldo Costo Total",
            "Descripción",
            "Recomendación"
        ]);

        // Anchos
        worksheet.getColumn(1).width = 18;
        worksheet.getColumn(2).width = 40;
        worksheet.getColumn(3).width = 10;
        worksheet.getColumn(4).width = 15;
        worksheet.getColumn(5).width = 20;
        worksheet.getColumn(6).width = 25;
        worksheet.getColumn(7).width = 18;
        worksheet.getColumn(8).width = 20;
        worksheet.getColumn(9).width = 20;
        worksheet.getColumn(10).width = 18;
        worksheet.getColumn(11).width = 20;
        worksheet.getColumn(12).width = 20;
        worksheet.getColumn(13).width = 55;
        worksheet.getColumn(14).width = 55;

        // Datos
        for (const result of results) {

            const metadata = result.metadata as Rule009Metadata;

            worksheet.addRow([
                result.product_code,
                result.product_name,
                metadata.month,
                metadata.date,
                metadata.document,
                metadata.operation,
                metadata.entryQuantity,
                metadata.entryUnitCost,
                metadata.entryTotalCost,
                metadata.balanceQuantity,
                metadata.balanceUnitCost,
                metadata.balanceTotalCost,
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