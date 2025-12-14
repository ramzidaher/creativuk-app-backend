import { Module } from '@nestjs/common';
import { OneDriveFileManagerService } from './onedrive-file-manager.service';
import { OneDriveFileManagerController } from './onedrive-file-manager.controller';
import { ContractModule } from '../contracts/contract.module';

@Module({
  imports: [ContractModule],
  providers: [OneDriveFileManagerService],
  controllers: [OneDriveFileManagerController],
  exports: [OneDriveFileManagerService],
})
export class OneDriveModule {}

