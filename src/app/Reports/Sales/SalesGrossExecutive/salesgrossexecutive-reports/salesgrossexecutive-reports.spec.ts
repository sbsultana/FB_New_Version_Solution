import { ComponentFixture, TestBed } from '@angular/core/testing';

import { SalesgrossexecutiveReports } from './salesgrossexecutive-reports';

describe('SalesgrossexecutiveReports', () => {
  let component: SalesgrossexecutiveReports;
  let fixture: ComponentFixture<SalesgrossexecutiveReports>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [SalesgrossexecutiveReports]
    })
    .compileComponents();

    fixture = TestBed.createComponent(SalesgrossexecutiveReports);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
