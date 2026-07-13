import ExcelJS from "exceljs";

export interface Rule012Metadata {
    issueDate: string;
    warehouseDate: string;
    supplierRuc: string;
    supplier: string;
    document: string;
    normalizedDocument: string;
    transitAmount: number;
    kardexAmount: number;
    difference: number;
    movements: number;
    transitItem: string;
}

export class Rule012Exporter {

    async export(results: any[]) {

        const workbook = new ExcelJS.Workbook();

        workbook.creator = "HM Kardex Audit";
        workbook.created = new Date();

        const worksheet = workbook.addWorksheet("RULE_012");

        worksheet.mergeCells("A1:M1");

        worksheet.getCell("A1").value =
            "RULE_012 - Validación del Monto de Mercadería en Tránsito vs Kardex";

        worksheet.getCell("A1").font = {
            bold: true,
            size: 16
        };

        // Espacio
        worksheet.addRow([]);

        // Cabeceras
        worksheet.addRow([
            "Fecha de Emisión",
            "Fecha de Ingreso a Almacén",
            "RUC Proveedor",
            "Proveedor",
            "Documento",
            "Documento Normalizado",
            "Monto Operación",
            "Monto Registrado en Kardex",
            "Diferencia",
            "Movimientos Encontrados",
            "Ítem Tránsito",
            "Descripción",
            "Recomendación"
        ]);

        // Anchos
        worksheet.getColumn(1).width = 18;
        worksheet.getColumn(2).width = 22;
        worksheet.getColumn(3).width = 18;
        worksheet.getColumn(4).width = 35;
        worksheet.getColumn(5).width = 25;
        worksheet.getColumn(6).width = 25;
        worksheet.getColumn(7).width = 22;
        worksheet.getColumn(8).width = 28;
        worksheet.getColumn(9).width = 18;
        worksheet.getColumn(10).width = 25;
        worksheet.getColumn(11).width = 18;
        worksheet.getColumn(12).width = 55;
        worksheet.getColumn(13).width = 55;

        // Datos
        for (const result of results) {

            const metadata = result.metadata as Rule012Metadata;

            worksheet.addRow([
                metadata.issueDate,
                metadata.warehouseDate,
                metadata.supplierRuc,
                metadata.supplier,
                metadata.document,
                metadata.normalizedDocument,
                metadata.transitAmount,
                metadata.kardexAmount,
                metadata.difference,
                metadata.movements,
                metadata.transitItem,
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
            to: "M3"
        };

        return workbook;
    }
}