import { ComponentFixture, TestBed } from '@angular/core/testing';

import { SalesgrossmanagerReports } from './salesgrossmanager-reports';

describe('SalesgrossmanagerReports', () => {
  let component: SalesgrossmanagerReports;
  let fixture: ComponentFixture<SalesgrossmanagerReports>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [SalesgrossmanagerReports]
    })
    .compileComponents();

    fixture = TestBed.createComponent(SalesgrossmanagerReports);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
