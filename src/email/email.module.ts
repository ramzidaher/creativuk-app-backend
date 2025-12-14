import { Module } from '@nestjs/common';
import { ConfigModule } from '@nestjs/config';
import { EmailService } from './email.service';
import { EmailController } from './email.controller';
import { CloudinaryService } from './cloudinary.service';

@Module({
  imports: [ConfigModule],
  providers: [EmailService, CloudinaryService],
  controllers: [EmailController],
  exports: [EmailService, CloudinaryService],
})
export class EmailModule {}