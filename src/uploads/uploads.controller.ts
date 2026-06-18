import { Controller, Get, Post, Body, Patch, Param, Delete } from '@nestjs/common';
import { UploadsService } from './uploads.service';
import { CreateUploadDto } from './dto/create-uploads.dto';
import { GenerateUploadUrlDto } from './dto/generate-uploads.dto';
import {
  UseGuards,
} from '@nestjs/common';
import {
  JwtAuthGuard,
} from '../auth/guards/jwt-auth.guard';
import { CurrentUser } from 'src/auth/decorators/current-user.decorator';

@UseGuards(JwtAuthGuard)
@Controller('uploads')
export class UploadsController {
    constructor(private readonly uploadsService: UploadsService) {}

    @Post()
    create(
        @Body()
        dto: CreateUploadDto,
        @CurrentUser()
        user: any,
        ) {
        return this.uploadsService.create(
            dto,
            user,
        );
    }

    @Get(':id/view')
    view(
    @Param('id')
    id: string,
    ) {
    return this.uploadsService.getViewUrl(
        id,
    );
    }

    @Get('audit/:auditId')
    findByAudit(
    @Param('auditId')
    auditId: string,
    ) {
    return this.uploadsService.findByAudit(
        auditId,
    );
    }

    @Delete(':id')
    remove(
    @Param('id')
    id: string,
    ) {
    return this.uploadsService.remove(
        id,
    );
    }    
    @Post('presigned-url')
    generateUploadUrl(
    @Body()
    dto: GenerateUploadUrlDto,
    ) {

    return this.uploadsService
        .generateUploadUrl(dto);

    }
}
