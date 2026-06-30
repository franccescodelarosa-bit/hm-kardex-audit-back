import { Injectable } from '@nestjs/common';
import { Prisma, PrismaClient } from '@prisma/client';
import { PrismaService } from '../../prisma/prisma.service';
import { FindingsQueryDto } from '../dto/findings-query.dto';

@Injectable()
export class AuditResultsRepository {
    constructor(
        private readonly prisma: PrismaService,
    ) {}
    async getAudits() {
        const audits = await this.prisma.audit_jobs.findMany({
            include:{
                clients:true
            },
            orderBy:{
                created_at:'desc'
            }
        });
        return audits.map(audit=>{

            return{
                id:audit.id,
                client: {
                    id: audit.clients.id,
                    businessName: audit.clients.business_name,
                    ruc: audit.clients.ruc
                },
                year:audit.year,
                status:audit.status,
                createdAt:audit.created_at,
                completedAt:audit.completed_at
            };

        });
    }
    async getDashboard(auditJobId: string) {
        const audit = await this.prisma.audit_jobs.findUnique({

            where: {
                id: auditJobId
            },

            include: {

                clients: true,

                audit_results: {

                    include: {

                        audit_rules: true

                    }

                }

            }

        });

        if (!audit) {
            return null;
        }

        const high = audit.audit_results.filter(x => x.audit_rules?.risk_level === 'HIGH').length;
        const medium = audit.audit_results.filter(x => x.audit_rules?.risk_level === 'MEDIUM').length;
        const low = audit.audit_results.filter(x => x.audit_rules?.risk_level === 'LOW').length;

        const topRules = new Map<string, any>();

        for (const finding of audit.audit_results) {

            if (!finding.audit_rules) continue;

            const id = finding.audit_rules.id;

            if (!topRules.has(id)) {

                topRules.set(id, {

                    id,

                    code: finding.audit_rules.code,

                    name: finding.audit_rules.name,

                    riskLevel: finding.audit_rules.risk_level,

                    count: 0

                });

            }

            topRules.get(id).count++;

        }

        return {

            audit: {

                id: audit.id,

                client: audit.clients.business_name,

                year: audit.year,

                status: audit.status,

                createdAt: audit.created_at,

                completedAt: audit.completed_at

            },

            summary: {

                totalFindings: audit.audit_results.length

            },

            riskLevels: {

                high,

                medium,

                low

            },

            topRules: Array
                .from(topRules.values())
                .sort((a, b) => b.count - a.count)

        };

    }
    async getRules(auditJobId: string) {
        const results = await this.prisma.audit_results.groupBy({

            by: ['rule_id'],

            where: {
                audit_job_id: auditJobId,
                rule_id: {
                    not: null
                }
            },

            _count: true

        });
        const rules = await this.prisma.audit_rules.findMany();

        return results.map(result => {
            const rule = rules.find(r => r.id === result.rule_id);
            return {

                ruleId: rule?.id,

                code: rule?.code,

                name: rule?.name,

                riskLevel: rule?.risk_level,

                findings: result._count

            };
        }).sort((a, b) => b.findings - a.findings);
    }
    async getFindings(
        auditJobId: string,
        dto: FindingsQueryDto
    ) {

        const where: Prisma.audit_resultsWhereInput = {

            audit_job_id: auditJobId

        };

        if (dto.ruleId) {

            where.rule_id = dto.ruleId;

        }

        if (dto.month) {

            where.month = dto.month;

        }

        if (dto.productCode) {

            where.product_code = {

                contains: dto.productCode,
                mode: 'insensitive'

            };

        }

        if (dto.search) {

            where.OR = [

                {

                    product_code: {

                        contains: dto.search,
                        mode: 'insensitive'

                    }

                },

                {

                    product_name: {

                        contains: dto.search,
                        mode: 'insensitive'

                    }

                },

                {

                    description: {

                        contains: dto.search,
                        mode: 'insensitive'

                    }

                }

            ];

        }

        const skip = (dto.page - 1) * dto.pageSize;

        const [items, total] = await Promise.all([

            this.prisma.audit_results.findMany({

                where,

                include: {

                    audit_rules: true

                },

                skip: (Number(dto.page) - 1) * Number(dto.pageSize),

                take: Number(dto.pageSize),

                orderBy: {

                    created_at: 'desc'

                }

            }),

            this.prisma.audit_results.count({

                where

            })

        ]);

        return {

            page: dto.page,

            pageSize: dto.pageSize,

            total,

            items: items.map(x => ({

                id: x.id,

                month: x.month,

                productCode: x.product_code,

                productName: x.product_name,

                description: x.description,

                recommendation: x.recommendation,

                rule: x.audit_rules ? {

                    id: x.audit_rules.id,

                    code: x.audit_rules.code,

                    name: x.audit_rules.name,

                    riskLevel: x.audit_rules.risk_level

                } : null

            }))

        };

    }
    async getFinding(id: string) {

        const finding = await this.prisma.audit_results.findUnique({

            where: {

                id

            },

            include: {

                audit_rules: true

            }

        });

        if (!finding) {

            return null;

        }

        return {

            id: finding.id,

            month: finding.month,

            productCode: finding.product_code,

            productName: finding.product_name,

            errorType: finding.error_type,

            description: finding.description,

            recommendation: finding.recommendation,

            metadata: finding.metadata,

            rule: finding.audit_rules ? {

                id: finding.audit_rules.id,

                code: finding.audit_rules.code,

                name: finding.audit_rules.name,

                riskLevel: finding.audit_rules.risk_level

            } : null

        };

    }
    
}