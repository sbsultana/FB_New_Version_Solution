import { ComponentFixture, TestBed } from '@angular/core/testing';

import { OPCodeTrackerDetails } from './opcode-tracker-details';

describe('OPCodeTrackerDetails', () => {
  let component: OPCodeTrackerDetails;
  let fixture: ComponentFixture<OPCodeTrackerDetails>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [OPCodeTrackerDetails]
    })
    .compileComponents();

    fixture = TestBed.createComponent(OPCodeTrackerDetails);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
