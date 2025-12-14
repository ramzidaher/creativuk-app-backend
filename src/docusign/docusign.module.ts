import { Module } from '@nestjs/common';
import { ConfigModule } from '@nestjs/config';
import { DocuSignService } from './docusign.service';
import { DocuSignController } from './docusign.controller';

@Module({
  imports: [ConfigModule],
  providers: [DocuSignService],
  controllers: [DocuSignController],
  exports: [DocuSignService],
})
export class DocuSignModule {}
