import { Type } from 'class-transformer';
import { IsInt, IsOptional, IsString, IsUUID, Min } from 'class-validator';

export class FindingsQueryDto {

    @IsOptional()
    @IsUUID()
    ruleId?: string;

    @IsOptional()
    @IsString()
    productCode?: string;

    @IsOptional()
    @IsString()
    search?: string;

    @IsOptional()
    @Type(() => Number)
    month?: number;

    @IsOptional()
    @Type(() => Number)
    @IsInt()
    @Min(1)
    page = 1;

    @IsOptional()
    @Type(() => Number)
    @IsInt()
    pageSize = 50;

    @IsOptional()
    riskLevel?: string;

}