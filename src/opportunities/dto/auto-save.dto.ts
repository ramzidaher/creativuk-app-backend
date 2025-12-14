import { IsString, IsOptional, IsObject, IsNotEmpty, IsBoolean } from 'class-validator';

export class AutoSaveDto {
  @IsString()
  @IsNotEmpty()
  opportunityId: string;

  @IsString()
  @IsNotEmpty()
  fieldName: string;

  @IsOptional()
  fieldValue?: any;

  @IsOptional()
  @IsObject()
  pageData?: any;

  @IsOptional()
  @IsString()
  pageName?: string;

  @IsOptional()
  @IsBoolean()
  skipLastPageUpdate?: boolean;
}

export class AutoSaveImageDto {
  @IsString()
  @IsNotEmpty()
  opportunityId: string;

  @IsString()
  @IsNotEmpty()
  fieldName: string;

  @IsString()
  @IsNotEmpty()
  base64Data: string;

  @IsString()
  @IsOptional()
  fileName?: string;

  @IsString()
  @IsOptional()
  mimeType?: string;

  @IsOptional()
  fileSize?: number;
}

export class GetAutoSaveDataDto {
  @IsString()
  @IsNotEmpty()
  opportunityId: string;

  @IsOptional()
  @IsString()
  pageName?: string;
}
