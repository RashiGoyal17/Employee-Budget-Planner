import { TestBed } from '@angular/core/testing';

import { BudgetPlanService } from './budget-plan-service';

describe('BudgetPlanService', () => {
  let service: BudgetPlanService;

  beforeEach(() => {
    TestBed.configureTestingModule({});
    service = TestBed.inject(BudgetPlanService);
  });

  it('should be created', () => {
    expect(service).toBeTruthy();
  });
});
