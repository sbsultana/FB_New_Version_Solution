import { ComponentFixture, TestBed } from '@angular/core/testing';

import { PartsGrossDetails } from './parts-gross-details';

describe('PartsGrossDetails', () => {
  let component: PartsGrossDetails;
  let fixture: ComponentFixture<PartsGrossDetails>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [PartsGrossDetails]
    })
    .compileComponents();

    fixture = TestBed.createComponent(PartsGrossDetails);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
