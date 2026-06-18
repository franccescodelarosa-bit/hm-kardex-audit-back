export class CreateUploadDto {

  auditJobId!: string;

  fileType!: string;

  month?: number;

  fileName!: string;

  s3Key!: string;

  fileSize?: number;

}