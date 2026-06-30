import ExcelJS from "exceljs";

export interface Rule007Metadata {
    month: number;
    movement: number;
    operation: string;
    document: string;

    previousBalance: {
        quantity: number;
        totalCost: number;
    };

    movimientos: {
        entryQuantity: number;
        exitQuantity: number;
        entryTotalCost: number;
        exitTotalCost: number;
    };

    expectedBalance: {
        quantity: number;
        totalCost: number;
    };

    actualBalance: {
        quantity: number;
        totalCost: number;
    };

    differences: string[];
}

export class Rule007Exporter {

    async export(results: any[]) {

        const workbook = new ExcelJS.Workbook();

        workbook.creator = "HM Kardex Audit";
        workbook.created = new Date();

        const worksheet = workbook.addWorksheet("RULE_007");

        worksheet.mergeCells("A1:S1");

        worksheet.getCell("A1").value =
            "RULE_007 - Validación de la Fórmula del Kardex";

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
            "Movimiento",
            "Operación",
            "Documento",
            "Saldo Anterior Cantidad",
            "Saldo Anterior Total",
            "Entrada Cantidad",
            "Salida Cantidad",
            "Entrada Total",
            "Salida Total",
            "Saldo Esperado Cantidad",
            "Saldo Real Cantidad",
            "Saldo Esperado Total",
            "Saldo Real Total",
            "Campos con Diferencia",
            "Descripción",
            "Recomendación"
        ]);

        // Anchos
        worksheet.getColumn(1).width = 18;
        worksheet.getColumn(2).width = 40;
        worksheet.getColumn(3).width = 10;
        worksheet.getColumn(4).width = 12;
        worksheet.getColumn(5).width = 20;
        worksheet.getColumn(6).width = 20;
        worksheet.getColumn(7).width = 20;
        worksheet.getColumn(8).width = 20;
        worksheet.getColumn(9).width = 18;
        worksheet.getColumn(10).width = 18;
        worksheet.getColumn(11).width = 18;
        worksheet.getColumn(12).width = 18;
        worksheet.getColumn(13).width = 20;
        worksheet.getColumn(14).width = 20;
        worksheet.getColumn(15).width = 20;
        worksheet.getColumn(16).width = 20;
        worksheet.getColumn(17).width = 35;
        worksheet.getColumn(18).width = 60;
        worksheet.getColumn(19).width = 60;

        // Datos
        for (const result of results) {

            const metadata = result.metadata as Rule007Metadata;

            worksheet.addRow([

                result.product_code,

                result.product_name,

                metadata.month,

                metadata.movement,

                metadata.operation,

                metadata.document,

                metadata.previousBalance.quantity,

                metadata.previousBalance.totalCost,

                metadata.movimientos.entryQuantity,

                metadata.movimientos.exitQuantity,

                metadata.movimientos.entryTotalCost,

                metadata.movimientos.exitTotalCost,

                metadata.expectedBalance.quantity,

                metadata.actualBalance.quantity,

                metadata.expectedBalance.totalCost,

                metadata.actualBalance.totalCost,

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
            to: "S3"
        };

        return workbook;

    }

}