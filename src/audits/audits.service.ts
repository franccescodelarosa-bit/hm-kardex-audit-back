import { Injectable } from '@nestjs/common';

import { PrismaService }
from '../prisma/prisma.service';

import { CreateAuditDto }
from './dto/create-audit.dto';

@Injectable()
export class AuditsService {

  constructor(
    private readonly prisma: PrismaService,
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

async findAll() {

  const audits =
    await this.prisma.audit_jobs.findMany({

      include: {
        clients: true,
        uploaded_files: {
          select: {
            id: true,
          },
        },
      },

      orderBy: {
        created_at: 'desc',
      },

    });

  const TOTAL_REQUIRED = 15;

  return audits.map(audit => {

    const uploadedCount = (audit.uploaded_files.length > TOTAL_REQUIRED) ? TOTAL_REQUIRED : audit.uploaded_files.length;

    const progress =
      Math.round(
        (uploadedCount / TOTAL_REQUIRED) * 100
      );

    return {
      ...audit,

      progress,

      uploadedCount,

      uploaded_files: undefined,
    };

  });

}

async findOne(id: string) {
    
  const audit =
    await this.prisma.audit_jobs.findUnique({
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

  const totalRequired = 15;

  const uploadedCount =
    audit.uploaded_files.length;

  let status = 'PENDING';

  if (uploadedCount > 0) {
    status = 'IN_PROGRESS';
  }

  if (uploadedCount >= totalRequired) {
    status = 'READY_FOR_AUDIT';
  }

  return JSON.parse(
    JSON.stringify(
      {
        ...audit,
        status,
        completion:
          Math.round(
            (
              uploadedCount /
              totalRequired
            ) * 100,
          ),
        uploadedCount,
        totalRequired,
      },
      (_, value) =>
        typeof value === 'bigint'
          ? Number(value)
          : value,
    ),
  );

}

}