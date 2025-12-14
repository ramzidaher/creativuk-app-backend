import { SetMetadata } from '@nestjs/common';
import { UserRole } from './dto/auth.dto';

export const Roles = (...roles: UserRole[]) => SetMetadata('roles', roles); 