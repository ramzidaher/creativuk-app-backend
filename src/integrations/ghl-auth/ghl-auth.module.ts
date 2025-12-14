import { Module } from '@nestjs/common';
import { GhlAuthController } from './ghl-auth.controller';
import { GhlAuthService } from './ghl-auth.service';
import { UserModule } from '../../user/user.module';
import { JwtModule } from '@nestjs/jwt';

@Module({
  imports: [
    UserModule,
    JwtModule.register({ secret: process.env.JWT_SECRET }),
  ],
  controllers: [GhlAuthController],
  providers: [GhlAuthService],
  exports: [GhlAuthService],
})
export class GhlAuthModule {} 