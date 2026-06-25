import { Injectable } from '@nestjs/common';
import { PrismaService } from '../prisma/prisma.service';
import { CreateAuditDto } from './dto/create-audit.dto';
import { AuditEngineService } from '../audit-engine/audit-engine.service';

@Injectable()
export class AuditsService {

  constructor(
    private readonly prisma: PrismaService,
    private readonly auditEngineService: AuditEngineService
  ) {}

  async create(
    createAuditDto: CreateAuditDto,
  ) {

    return this.prisma.audit_jobs.create({
      data: {

        client_id:
          createAuditDto.clientId,

        year:
          createAuditDto.year,

        status:
          'PENDING',

        total_errors:
          0,

      },
    });

  }

  private readonly TOTAL_REQUIRED = 15;

  private calculateAuditStatus(
    uploadedCount: number,
  ) {

    if (
      uploadedCount >=
      this.TOTAL_REQUIRED
    ) {
      return 'READY_FOR_AUDIT';
    }

    if (uploadedCount > 0) {
      return 'IN_PROGRESS';
    }

    return 'PENDING';

  }

  private calculateProgress( uploadedCount: number ) {
    const files = Math.min( uploadedCount, this.TOTAL_REQUIRED );
    return {
      uploadedCount: files,
      progress: Math.round(
        (
          files /
          this.TOTAL_REQUIRED
        ) * 100,
      )
    };

  }

  async findAll() {
    const audits = await this.prisma.audit_jobs.findMany({
      include: {
        clients: true,
        _count: {
          select: {
            uploaded_files: true,
          },
        },
      },
      orderBy: {
        created_at: 'desc',
      },
    });
    return audits.map(audit => {
      const meta = this.calculateProgress( audit._count.uploaded_files);
      return {
        ...audit,
        status: audit.status,
        progress: meta.progress,
        uploadedCount: meta.uploadedCount,
      };
    });
  }

  async findOne( id: string) {
    const audit = await this.prisma.audit_jobs.findUnique({
        where: {
          id,
        },
        include: {
          clients: true,
          uploaded_files: true,
        },
    });

    if (!audit) {
      return null;
    }

    const meta =
      this.calculateProgress(
        audit.uploaded_files.length,
      );

    return JSON.parse(
      JSON.stringify(
        {

          ...audit,

          status: audit.status,
          completion: meta.progress,
          uploadedCount: meta.uploadedCount,
          totalRequired: this.TOTAL_REQUIRED,
        },

        (_, value) =>
          typeof value ===
          'bigint'
            ? Number(value)
            : value,

      ),
    );

  }

async run( id: string ) {
  const audit = await this.prisma.audit_jobs.findUnique({
    where: {
      id,
    },
  });
  if (!audit) {
    throw new Error(
      'Auditoría no encontrada',
    );
  }
  await this.prisma.audit_jobs.update({
    where: {
      id,
    },
    data: {
      status: 'AUDITING',
      started_at: new Date(),
    },
  });
  await this.prisma.audit_job_logs.create({
    data: {
      audit_job_id: id,
      step: 'AUDIT_STARTED',
      status: 'RUNNING',
      message:
        'Proceso de auditoría iniciado',
    },
  });
  await this.auditEngineService.enqueueAudit(id);

  return {
    success: true,
    message:
      'Auditoría enviada a procesamiento',
  };

}

}