import { Injectable } from '@nestjs/common';
import { PrismaService }
from '../prisma/prisma.service';

import { CreateUploadDto }
from './dto/create-uploads.dto';

import { GenerateUploadUrlDto }
from './dto/generate-uploads.dto';

import {
  S3Client,
  PutObjectCommand,
} from '@aws-sdk/client-s3';

import {
  getSignedUrl,
} from '@aws-sdk/s3-request-presigner';

import {
  GetObjectCommand,
} from '@aws-sdk/client-s3';

@Injectable()
export class UploadsService {

  constructor(
    private readonly prisma: PrismaService,
  ) {}

  private readonly s3 =
  new S3Client({
    region:
      process.env.AWS_REGION,
  });

  async generateUploadUrl(
    dto: GenerateUploadUrlDto,
    ) {

        const audit =
        await this.prisma.audit_jobs.findUnique({
            where: {
            id: dto.auditJobId,
            },
        });

        if (!audit) {
        throw new Error(
            'Auditoría no encontrada',
        );
        }

    //const key = `clients/${dto.auditJobId}/${dto.fileName}`;
    let key = '';
    if (
        dto.fileType ===
        'KARDEX'
        ) {
        key =
            `audits/${audit.year}/${dto.auditJobId}/KARDEX/${String(dto.month).padStart(2, '0')}/${dto.fileName}`;

        } else {
        key =
            `audits/${audit.year}/${dto.auditJobId}/${dto.fileType}/${dto.fileName}`;

        }

    const command =
        new PutObjectCommand({

        Bucket:
            process.env.AWS_S3_BUCKET,

        Key:
            key,

        });

    const uploadUrl =
        await getSignedUrl(
        this.s3,
        command,
        {
            expiresIn: 300,
        },
        );

    return {
        uploadUrl,
        key,
    };

    }

   async getViewUrl(
  id: string,
) {

  const upload =
    await this.prisma.uploaded_files.findUnique({
      where: {
        id,
      },
    });

  if (!upload) {
    throw new Error(
      'Archivo no encontrado',
    );
  }

  const command =
    new GetObjectCommand({
      Bucket:
        process.env.AWS_S3_BUCKET,
      Key:
        upload.s3_key,
    });

  const url =
    await getSignedUrl(
      this.s3,
      command,
      {
        expiresIn: 300,
      },
    );

  return {
    url,
  };
   }

  async create(dto: CreateUploadDto, user: any) {
    const dbUser =
    await this.prisma.users.findFirst({
        where: {
        cognito_sub: user.sub,
        },
    });

    if (!dbUser) {
    throw new Error(
        'Usuario no encontrado',
    );
    }
    const upload =
        await this.prisma.uploaded_files.create({
        data: {
            audit_job_id: dto.auditJobId,
            file_type: dto.fileType,
            month: dto.month,
            file_name: dto.fileName,
            s3_key: dto.s3Key,
            file_size: dto.fileSize,
            status: 'UPLOADED',
            uploaded_by: dbUser.id
        },
        });

    return {
        ...upload,
        file_size:
        upload.file_size
            ? Number(upload.file_size)
            : null,
    };

    }

    async findByAudit(
        auditJobId: string,
        ) {

        const uploads =
            await this.prisma.uploaded_files.findMany({
            where: {
                audit_job_id: auditJobId,
            },
            include: {
                users: {
                select: {
                    full_name: true,
                },
                },
            },
            orderBy: {
                uploaded_at: 'desc',
            },
            });

        return uploads.map(upload => ({
            id: upload.id,
            audit_job_id: upload.audit_job_id,
            file_type: upload.file_type,
            month: upload.month,
            file_name: upload.file_name,
            s3_key: upload.s3_key,
            status: upload.status,
            uploaded_at: upload.uploaded_at,

            file_size:
            upload.file_size
                ? Number(upload.file_size)
                : null,

            uploadedBy:
            upload.users?.full_name ??
            'Usuario desconocido',
        }));
    }

  async remove(
    id: string,
  ) {

    return this.prisma.uploaded_files.delete({
      where: {
        id,
      },
    });

  }

}
