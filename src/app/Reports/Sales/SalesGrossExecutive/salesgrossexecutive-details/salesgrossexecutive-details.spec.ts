import { ComponentFixture, TestBed } from '@angular/core/testing';

import { SalesgrossexecutiveDetails } from './salesgrossexecutive-details';

describe('SalesgrossexecutiveDetails', () => {
  let component: SalesgrossexecutiveDetails;
  let fixture: ComponentFixture<SalesgrossexecutiveDetails>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [SalesgrossexecutiveDetails]
    })
    .compileComponents();

    fixture = TestBed.createComponent(SalesgrossexecutiveDetails);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
