import { Injectable } from '@nestjs/common';
import { Prisma, PrismaClient } from '@prisma/client';
import { PrismaService } from '../../prisma/prisma.service';
import { FindingsQueryDto } from '../dto/findings-query.dto';
import { UpdateAuditFollowUpDto } from '../dto/update-audit.dto';

@Injectable()
export class AuditResultsRepository {
    constructor(
        private readonly prisma: PrismaService,
    ) {}
    async getAudits() {
        const audits = await this.prisma.audit_jobs.findMany({
            include: {
                clients: true
            },
            orderBy: {
                created_at: "desc"
            }
        });
        return audits.map(audit => {
            return {
                id: audit.id,
                client: {
                    id: audit.clients.id,
                    businessName: audit.clients.business_name,
                    ruc: audit.clients.ruc
                },
                year: audit.year,
                status: audit.status,
                createdAt: audit.created_at,
                completedAt: audit.completed_at,
                validationStatus: audit.validation_status,
                responsible: audit.responsible,
                regularizationDate: audit.regularization_date,
                correctiveAction: audit.corrective_action,
                observations: audit.observations
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
        const high = audit.audit_results.filter(x => x.audit_rules?.risk_level === 'CRITICO').length;
        const medium = audit.audit_results.filter(x => x.audit_rules?.risk_level === 'ALTO').length;
        const low = audit.audit_results.filter(x => x.audit_rules?.risk_level === 'MEDIO').length;
        const topRules = new Map<string, any>();
        const TOTAL_RULES = 14;
        const affectedProducts = new Set(
            audit.audit_results
                .map(x => x.product_code)
                .filter(Boolean)
        ).size;
        let economicImpact = 0;
        let generalStatus = "CON_OBSERVACIONES";
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
            const metadata: any = finding.metadata ?? {};
            switch (finding.audit_rules?.code) {
                case "RULE_001":
                    economicImpact += Math.abs(
                        Number(metadata.inventoryTotalCost ?? 0) -
                        Number(metadata.kardexTotalCost ?? 0)
                    );
                    break;
                case "RULE_002":
                case "RULE_003":
                case "RULE_013":
                case "RULE_014":
                    economicImpact += Math.abs(
                        Number(
                            metadata?.difference?.totalCost ??
                            metadata?.difference ??
                            0
                        )
                    );
                    break;
                case "RULE_012":
                    economicImpact += Math.abs(
                        Number(metadata.difference ?? 0)
                    );
                    break;
            }
            economicImpact = Number(
                economicImpact.toFixed(2)
            );
        }        
        const failedRules = topRules.size;
        const passedRules = TOTAL_RULES - failedRules;
        const compliance = Number(
            ((passedRules / TOTAL_RULES) * 100).toFixed(2)
        );
        
            const hasCriticalRule = Array
                .from(topRules.values())
                .some(x =>
                    x.riskLevel === "CRITICO"
                );
                if (
                    hasCriticalRule ||
                    compliance < 80
                ) {
                    generalStatus = "CRITICO";
                }
                else if (
                    compliance >= 95 &&
                    economicImpact === 0
                ) {
                    generalStatus = "APROBADO";
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
                totalFindings: audit.audit_results.length,
                executedRules: TOTAL_RULES,
                passedRules,
                failedRules,
                compliance,
                affectedProducts,
                economicImpact,
                generalStatus
            },
            followUp: {
                validationStatus: audit.validation_status,
                regularizationDate: audit.regularization_date,
                correctiveAction: audit.corrective_action,
                observations: audit.observations,
                responsible: audit.responsible
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
        const [results, rules] = await Promise.all([
            this.prisma.audit_results.groupBy({
                by: ['rule_id'],
                where: {
                    audit_job_id: auditJobId,
                    rule_id: {
                        not: null
                    }
                },
                _count: {
                    _all: true
                }
            }),
            this.prisma.audit_rules.findMany()
        ]);

        return results
            .map(result => {
                const rule = rules.find(r => r.id === result.rule_id);
                return {
                    id: rule?.id,
                    code: rule?.code,
                    name: rule?.name,
                    riskLevel: rule?.risk_level,
                    count: result._count._all
                };
            })
            .sort((a, b) => (a.code ?? "").localeCompare(b.code ?? ""));
    }
    async getFindings(auditJobId: string, dto: FindingsQueryDto) {
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
        const [items, total, risks] = await Promise.all([
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
            }),
            this.prisma.audit_results.groupBy({
                by: ['risk_level'],
                where,
                _count: {
                    _all: true
                }
            })
        ]);
        const summary = { critical: 0, high: 0,medium: 0};
        for (const risk of risks) {
            switch (risk.risk_level?.toUpperCase()) {
                case "CRITICO":
                case "CRÍTICO":
                    summary.critical = risk._count._all;
                    break;
                case "ALTO":
                    summary.high = risk._count._all;
                    break;
                case "MEDIO":
                    summary.medium = risk._count._all;
                    break;
            }
        }
        return {
            page: dto.page,
            pageSize: dto.pageSize,
            total,
            summary,
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
    async putAuditResult(id: string, dto: UpdateAuditFollowUpDto ) {
        return this.prisma.audit_jobs.update({
            where: {
                id
            },
            data: {
                validation_status: dto.validationStatus,
                regularization_date: dto.regularizationDate
                ? new Date(dto.regularizationDate)
                : null,
                corrective_action: dto.correctiveAction,
                observations: dto.observations,
                responsible: dto.responsible
            }
        });
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
    async getExcel(auditJobId: string, ruleId: string) {
        const auditJob = await this.prisma.audit_jobs.findUnique({
            where: {
                id: auditJobId
            },
            select: {
                year: true,
                clients: {
                    select: {
                        business_name: true,
                        ruc: true
                    }
                }
            }
        });
        const rows = await this.prisma.audit_results.findMany({
            where: {
                audit_job_id: auditJobId,
                rule_id: ruleId
            },
            select: {
                month: true,
                product_code: true,
                product_name: true,
                description: true,
                recommendation: true,
                error_type: true,
                risk_level: true,
                metadata: true,
                audit_rules: {
                    select: {
                        id: true,
                        code: true,
                        name: true
                    }
                }
            },
            orderBy: [
                {
                    month: "asc"
                },
                {
                    product_code: "asc"
                }
            ]
        });
        console.log(rows);
        if (!rows.length) {
            return {
                header: null,
                rows: []
            };
        }
        const first = rows[0];
        return {
            header: {
                companyName: auditJob!.clients.business_name,
                ruc: auditJob!.clients.ruc ?? "",
                year: auditJob!.year
            },
            rows
        };
    }
}