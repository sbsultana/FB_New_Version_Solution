import { ComponentFixture, TestBed } from '@angular/core/testing';

import { ServiceOpenRODetails } from './service-open-rodetails';

describe('ServiceOpenRODetails', () => {
  let component: ServiceOpenRODetails;
  let fixture: ComponentFixture<ServiceOpenRODetails>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [ServiceOpenRODetails]
    })
    .compileComponents();

    fixture = TestBed.createComponent(ServiceOpenRODetails);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
