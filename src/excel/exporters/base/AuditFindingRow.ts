export interface AuditFindingRow {

    period: string;

    productCode: string;

    productDescription: string;

    inconsistencyType: string;

    expectedValue: string | number;

    foundValue: string | number;

    difference: string | number;

    differencePercent?: number;

    riskLevel: string;

    traceability: string;

}