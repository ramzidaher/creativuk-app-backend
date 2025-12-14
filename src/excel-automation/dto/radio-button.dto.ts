import { IsString, IsNotEmpty, IsOptional } from 'class-validator';

export class SelectRadioButtonDto {
  @IsString()
  @IsNotEmpty()
  shapeName: string;

  @IsString()
  @IsOptional()
  opportunityId?: string;
}

export class RadioButtonResponseDto {
  success: boolean;
  message: string;
  shapeName?: string;
  error?: string;
}
