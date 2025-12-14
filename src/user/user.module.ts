import { Module } from '@nestjs/common';
import { UserService } from './user.service';
import { UserController } from './user.controller';
import { PrismaModule } from '../prisma/prisma.module';
import { GoHighLevelService } from '../integrations/gohighlevel.service';
import { GHLUserLookupService } from '../integrations/ghl-user-lookup.service';
import { PrismaService } from '../prisma/prisma.service';

@Module({
  imports: [PrismaModule],
  controllers: [UserController],
  providers: [UserService, GoHighLevelService, GHLUserLookupService, PrismaService],
  exports: [UserService],
})
export class UserModule {}
