import { Injectable, BadRequestException } from '@nestjs/common';
import { AuditResultsRepository } from './repositories/audit-results.repository';
import { FindingsQueryDto } from './dto/findings-query.dto';
import { Rule001Exporter } from '../excel/exporters/Rule001Exporter';
import { Rule002Exporter } from '../excel/exporters/Rule002Exporter';
import { Rule003Exporter } from '../excel/exporters/Rule003Exporter';
import { Rule004Exporter } from '../excel/exporters/Rule004Exporter';
import { Rule005Exporter } from '../excel/exporters/Rule005Exporter';
import { Rule006Exporter } from '../excel/exporters/Rule006Exporter';
import { Rule007Exporter } from '../excel/exporters/Rule007Exporter';
import { Rule008Exporter } from '../excel/exporters/Rule008Exporter';
import { Rule009Exporter } from '../excel/exporters/Rule009Exporter';
import { Rule010Exporter } from '../excel/exporters/Rule010Exporter';
import { Rule011Exporter } from '../excel/exporters/Rule011Exporter';
import { Rule012Exporter } from '../excel/exporters/Rule012Exporter';
import { Rule013Exporter } from '../excel/exporters/Rule013Exporter';
import { Rule014Exporter } from '../excel/exporters/Rule014Exporter';


@Injectable()
export class AuditResultsService {
    constructor(
        private readonly repository: AuditResultsRepository
    ){}
    async getAudits(){
        return this.repository.getAudits();
    }
    async getDashboard(auditJobId: string) {
        return this.repository.getDashboard(auditJobId);
    }
    async getRules(auditJobId: string) {
        return this.repository.getRules(auditJobId);
    }
    async getFindings(
        auditJobId: string,
        dto: FindingsQueryDto
    ) {

        return this.repository.getFindings(auditJobId, dto);

    }
    async getFinding(id: string) {
        return this.repository.getFinding(id);
    }
    async getExcel(auditJobId: string, ruleId: string){
        const rows = await this.repository.getExcel(auditJobId, ruleId);
        if (!rows.length) {
            throw new BadRequestException("No existen resultados para la regla.");
        }
        switch(rows[0].audit_rules!.code){
            case "RULE_001":
                const exporter = new Rule001Exporter();
                return exporter.export(rows);                
            case "RULE_002":
                const exporter2 = new Rule002Exporter();
                return exporter2.export(rows);
            case "RULE_003":
                const exporter3 = new Rule003Exporter();
                return exporter3.export(rows);
            case "RULE_004":
                const exporter4 = new Rule004Exporter();
                return exporter4.export(rows);
            case "RULE_005":
                const exporter5 = new Rule005Exporter();
                return exporter5.export(rows);
            case "RULE_006":
                const exporter6 = new Rule006Exporter();
                return exporter6.export(rows);
            case "RULE_007":
                const exporter7 = new Rule007Exporter();
                return exporter7.export(rows);
            case "RULE_008":
                const exporter8 = new Rule008Exporter();
                return exporter8.export(rows);
            case "RULE_009":
                const exporter9 = new Rule009Exporter();
                return exporter9.export(rows);
            case "RULE_010":
                const exporter10 = new Rule010Exporter();
                return exporter10.export(rows);
            case "RULE_011":
                const exporter11 = new Rule011Exporter();
                return exporter11.export(rows); 
            case "RULE_012":
                const exporter12 = new Rule012Exporter();
                return exporter12.export(rows); 
            case "RULE_013":
                const exporter13 = new Rule013Exporter();
                return exporter13.export(rows); 
            case "RULE_014":
                const exporter14 = new Rule014Exporter();
                return exporter14.export(rows); 
            default:
                throw new BadRequestException(
                    `Exportador no implementado para la regla ${rows[0].audit_rules!.code}`
                );
        }
    }
}
