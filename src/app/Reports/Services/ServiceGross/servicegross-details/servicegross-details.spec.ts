import { ComponentFixture, TestBed } from '@angular/core/testing';

import { ServicegrossDetails } from './servicegross-details';

describe('ServicegrossDetails', () => {
  let component: ServicegrossDetails;
  let fixture: ComponentFixture<ServicegrossDetails>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [ServicegrossDetails]
    })
    .compileComponents();

    fixture = TestBed.createComponent(ServicegrossDetails);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
