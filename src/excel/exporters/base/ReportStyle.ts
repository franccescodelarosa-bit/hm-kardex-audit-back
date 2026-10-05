
export type Align = "center" | "general";

export interface ColumnStyle {
    width: number;
    /** Alineación horizontal de los datos. */
    align: Align;
    /** Tamaño de letra de los datos. */
    size: number;
    /** Tamaño de letra del encabezado de esta columna (por defecto, el del encabezado). */
    headerSize?: number;
    /** Formato numérico de los datos (ej. "#,##0.00"). */
    numFmt?: string;
}

export interface ReportStyle {
    titleSize: number;
    titleHeight: number;
    /** Tamaño de letra de la fila 2 (Empresa / RUC / Periodo). */
    infoSize: number;
    /** Alineación del nombre de la empresa (columna B de la fila 2). */
    companyAlign: Align;
    /** Tamaño de letra del nombre de la empresa, si difiere del resto de la fila 2. */
    companySize?: number;
    headerSize: number;
    headerHeight: number;
    columns: ColumnStyle[];
    subHeader?: {
        height: number;
        /** Columna (base 1) -> texto. */
        labels: Record<number, string>;
    };
}

export const MONEY_FORMAT = "#,##0.00";

interface ColumnsSpec {
    widths: number[];
    /** Un carácter por columna: "c" = centrado, cualquier otro = general. */
    align: string;
    size: number;
    /** Tamaños de letra distintos del general, por columna (base 1). */
    sizes?: Record<number, number>;
    /** Tamaños de letra del encabezado distintos del general, por columna (base 1). */
    headerSizes?: Record<number, number>;
    /** Columnas (base 1) con formato de monto. */
    money?: number[];
}

function columnsOf(spec: ColumnsSpec): ColumnStyle[] {
    return spec.widths.map((width, index) => {
        const column = index + 1;
        const style: ColumnStyle = {
            width,
            align: spec.align[index] === "c" ? "center" : "general",
            size: spec.sizes?.[column] ?? spec.size
        };
        if (spec.headerSizes?.[column] !== undefined) {
            style.headerSize = spec.headerSizes[column];
        }
        if (spec.money?.includes(column)) {
            style.numFmt = MONEY_FORMAT;
        }
        return style;
    });
}

export const DEFAULT_STYLE: ReportStyle = {
    titleSize: 18,
    titleHeight: 23.25,
    infoSize: 11,
    companyAlign: "center",
    headerSize: 11,
    headerHeight: 30,
    columns: columnsOf({
        widths: [14, 18, 40, 30, 20, 20, 18, 16, 16, 50],
        align: "-c--ccccc-",
        size: 11,
        money: [5, 6, 7]
    })
};

export const REPORT_STYLES: Record<string, ReportStyle> = {
    RULE_001: {
        titleSize: 18,
        titleHeight: 23.25,
        infoSize: 11,
        companyAlign: "center",
        headerSize: 11,
        headerHeight: 30,
        columns: columnsOf({
            widths: [14, 13.57, 22.71, 14, 15, 15, 12.14, 10.29, 11.86, 41.14],
            align: "-c--ccccc-",
            size: 11
        }),
        subHeader: {
            height: 45,
            labels: {
                5: "Valor esperado del inventario final",
                6: "Valor encontrado del Kardex"
            }
        }
    },
    RULE_002: {
        titleSize: 16,
        titleHeight: 21,
        infoSize: 11,
        companyAlign: "center",
        headerSize: 11,
        headerHeight: 45,
        columns: columnsOf({
            widths: [14, 11.29, 23.29, 16, 10.86, 10.86, 10.86, 10.86, 10.86, 33],
            align: "-c--ccccc-",
            size: 11
        })
    },
    RULE_003: {
        titleSize: 16,
        titleHeight: 21,
        infoSize: 11,
        companyAlign: "center",
        headerSize: 11,
        headerHeight: 30,
        columns: columnsOf({
            widths: [14, 18, 23.86, 16.71, 12, 11.57, 13.86, 10.57, 10.14, 34],
            align: "-c---cccc-",
            size: 11,
            money: [5]
        })
    },
    RULE_004: {
        titleSize: 18,
        titleHeight: 23.25,
        infoSize: 12,
        companyAlign: "center",
        headerSize: 12,
        headerHeight: 47.25,
        columns: columnsOf({
            widths: [14, 14, 16, 19.57, 28, 20, 21.71, 22, 59, 22, 18, 18, 16, 14, 16, 76],
            align: "-cc------c----c-",
            size: 12,
            sizes: { 9: 10, 16: 10 },
            headerSizes: { 9: 10, 16: 10 },
            money: [11, 12, 13]
        })
    },
    RULE_005: {
        titleSize: 16,
        titleHeight: 21,
        infoSize: 11,
        companyAlign: "general",
        headerSize: 11,
        headerHeight: 30,
        columns: columnsOf({
            widths: [14, 13.43, 25.43, 18.71, 12.86, 12.86, 12.86, 12.86, 12.86, 34],
            align: "-c-cccccc-",
            size: 14,
            sizes: { 10: 11 }
        })
    },
    RULE_006: {
        titleSize: 20,
        titleHeight: 26.25,
        infoSize: 12,
        companyAlign: "general",
        headerSize: 12,
        headerHeight: 31.5,
        columns: columnsOf({
            widths: [14, 12.43, 23.57, 15, 11, 12.86, 12.57, 12.14, 11, 34.14],
            align: "-c-cccccc-",
            size: 14,
            sizes: { 10: 11 }
        })
    },
    RULE_008: {
        titleSize: 20,
        titleHeight: 26.25,
        infoSize: 14,
        companyAlign: "general",
        headerSize: 14,
        headerHeight: 37.5,
        columns: columnsOf({
            widths: [14, 16.29, 34.29, 27, 14.43, 18.14, 15.43, 12.57, 14.14, 35.14],
            align: "-c--ccccc-",
            size: 14,
            sizes: { 10: 11 },
            headerSizes: { 10: 11 }
        })
    },
    RULE_011: {
        titleSize: 18,
        titleHeight: 23.25,
        infoSize: 12,
        companyAlign: "general",
        headerSize: 12,
        headerHeight: 31.5,
        columns: columnsOf({
            widths: [14, 18, 38, 21.29, 20, 20, 18, 16, 16, 30.57],
            align: "-c-c------",
            size: 14,
            sizes: { 1: 11, 10: 11 },
            money: [5, 6, 7]
        })
    },
    RULE_012: {
        titleSize: 20,
        titleHeight: 26.25,
        infoSize: 14,
        companyAlign: "general",
        headerSize: 12,
        headerHeight: 31.5,
        columns: columnsOf({
            widths: [14, 18, 21, 17.57, 15.14, 14.43, 15.14, 13.57, 11.57, 50],
            align: "-------cc-",
            size: 14,
            sizes: { 10: 12 },
            money: [5, 6, 7]
        })
    },
    RULE_013: {
        titleSize: 18,
        titleHeight: 23.25,
        infoSize: 14,
        companyAlign: "center",
        headerSize: 12,
        headerHeight: 31.5,
        columns: columnsOf({
            widths: [14, 14, 40, 15, 16.29, 16.29, 16.29, 12.57, 13, 50],
            align: "-c-c------",
            size: 14,
            sizes: { 10: 12 },
            money: [5, 6, 7]
        })
    },
    RULE_014: {
        titleSize: 18,
        titleHeight: 23.25,
        infoSize: 14,
        companyAlign: "general",
        companySize: 12,
        headerSize: 14,
        headerHeight: 31.5,
        columns: columnsOf({
            widths: [14, 18, 40, 30, 20, 20, 18, 16, 16, 50],
            align: "-------cc-",
            size: 14,
            sizes: { 2: 12, 10: 12 },
            headerSizes: { 2: 12, 10: 12 },
            money: [5, 6, 7]
        })
    }
};

export function styleFor(sheetName: string): ReportStyle {
    return REPORT_STYLES[sheetName] ?? DEFAULT_STYLE;
}
