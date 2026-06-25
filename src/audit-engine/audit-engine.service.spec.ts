import { Test, TestingModule } from '@nestjs/testing';
import { AuditEngineService } from './audit-engine.service';

describe('AuditEngineService', () => {
  let service: AuditEngineService;

  beforeEach(async () => {
    const module: TestingModule = await Test.createTestingModule({
      providers: [AuditEngineService],
    }).compile();

    service = module.get<AuditEngineService>(AuditEngineService);
  });

  it('should be defined', () => {
    expect(service).toBeDefined();
  });
});
