import { Module } from '@nestjs/common';
import { FreeSignaturesController } from './free-signatures.controller';
import { FreeSignaturesService } from './free-signatures.service';
import { PrismaModule } from '../prisma/prisma.module';

@Module({
  imports: [PrismaModule],
  controllers: [FreeSignaturesController],
  providers: [FreeSignaturesService],
  exports: [FreeSignaturesService],
})
export class FreeSignaturesModule {}
