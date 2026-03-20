import { ComponentFixture, TestBed } from '@angular/core/testing';

import { PartsOpenRODetails } from './parts-open-rodetails';

describe('PartsOpenRODetails', () => {
  let component: PartsOpenRODetails;
  let fixture: ComponentFixture<PartsOpenRODetails>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [PartsOpenRODetails]
    })
    .compileComponents();

    fixture = TestBed.createComponent(PartsOpenRODetails);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
