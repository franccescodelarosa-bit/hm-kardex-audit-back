import { Injectable } from '@nestjs/common';
import { AuditResultsRepository } from './repositories/audit-results.repository';
import { FindingsQueryDto } from './dto/findings-query.dto';

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
}
