import { ComponentFixture, TestBed } from '@angular/core/testing';

import { SalesgrossmanagerDetails } from './salesgrossmanager-details';

describe('SalesgrossmanagerDetails', () => {
  let component: SalesgrossmanagerDetails;
  let fixture: ComponentFixture<SalesgrossmanagerDetails>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [SalesgrossmanagerDetails]
    })
    .compileComponents();

    fixture = TestBed.createComponent(SalesgrossmanagerDetails);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
