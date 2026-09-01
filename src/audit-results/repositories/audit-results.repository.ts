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

    //todo: metadata de estas reglas
    private static readonly ECONOMIC_IMPACT_RULE_CODES = [
        "RULE_001",
        "RULE_002",
        "RULE_003",
        "RULE_012",
        "RULE_013",
        "RULE_014"
    ];

    private static computeEconomicImpact(
        rows: { metadata: Prisma.JsonValue; audit_rules: { code: string } | null }[]
    ): number {
        let economicImpact = 0;
        for (const row of rows) {
            const metadata: any = row.metadata ?? {};
            switch (row.audit_rules?.code) {
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
        }
        return Number(economicImpact.toFixed(2));
    }

    async getDashboard(auditJobId: string) {

        const TOTAL_RULES = 14;

        //todo: conteo y data puntual
        const [
            audit,
            totalFindings,
            ruleCounts,
            productGroups,
            economicRows,
            allRules
        ] = await Promise.all([
            this.prisma.audit_jobs.findUnique({
                where: { id: auditJobId },
                select: {
                    id: true,
                    year: true,
                    status: true,
                    created_at: true,
                    completed_at: true,
                    validation_status: true,
                    regularization_date: true,
                    corrective_action: true,
                    observations: true,
                    responsible: true,
                    clients: {
                        select: { business_name: true }
                    }
                }
            }),
            this.prisma.audit_results.count({
                where: { audit_job_id: auditJobId }
            }),
            this.prisma.audit_results.groupBy({
                by: ['rule_id'],
                where: { audit_job_id: auditJobId, rule_id: { not: null } },
                _count: { _all: true }
            }),
            this.prisma.audit_results.groupBy({
                by: ['product_code'],
                where: { audit_job_id: auditJobId, product_code: { not: null } }
            }),
            this.prisma.audit_results.findMany({
                where: {
                    audit_job_id: auditJobId,
                    audit_rules: {
                        code: { in: AuditResultsRepository.ECONOMIC_IMPACT_RULE_CODES }
                    }
                },
                select: {
                    metadata: true,
                    audit_rules: { select: { code: true } }
                }
            }),
            this.prisma.audit_rules.findMany()
        ]);

        if (!audit) {
            return null;
        }

        const rulesById = new Map(allRules.map(rule => [rule.id, rule]));
        const topRules = ruleCounts
            .map(group => {
                const rule = rulesById.get(group.rule_id!);
                return {
                    id: group.rule_id!,
                    code: rule?.code,
                    name: rule?.name,
                    riskLevel: rule?.risk_level,
                    count: group._count._all
                };
            })
            .sort((a, b) => b.count - a.count);

        //todo: contero de riesgos fijo de la regla
        let high = 0;
        let medium = 0;
        let low = 0;
        for (const rule of topRules) {
            if (rule.riskLevel === 'CRITICO') high += rule.count;
            else if (rule.riskLevel === 'ALTO') medium += rule.count;
            else if (rule.riskLevel === 'MEDIO') low += rule.count;
        }

        const affectedProducts = productGroups.filter(
            group => Boolean(group.product_code)
        ).length;

        const economicImpact =
            AuditResultsRepository.computeEconomicImpact(economicRows as any);

        const failedRules = topRules.length;
        const passedRules = TOTAL_RULES - failedRules;
        const compliance = Number(
            ((passedRules / TOTAL_RULES) * 100).toFixed(2)
        );

        const hasCriticalRule = topRules.some(x => x.riskLevel === "CRITICO");

        let generalStatus = "CON_OBSERVACIONES";
        if (hasCriticalRule || compliance < 80) {
            generalStatus = "CRITICO";
        } else if (compliance >= 95 && economicImpact === 0) {
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
                totalFindings,
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
            topRules
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
        if (dto.riskLevel) {
            where.risk_level = dto.riskLevel;
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
                riskLevel: x.risk_level,
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
            riskLevel: finding.risk_level,
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