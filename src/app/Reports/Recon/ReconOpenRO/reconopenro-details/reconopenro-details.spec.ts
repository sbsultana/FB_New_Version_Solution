import { ComponentFixture, TestBed } from '@angular/core/testing';

import { ReconopenroDetails } from './reconopenro-details';

describe('ReconopenroDetails', () => {
  let component: ReconopenroDetails;
  let fixture: ComponentFixture<ReconopenroDetails>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [ReconopenroDetails]
    })
    .compileComponents();

    fixture = TestBed.createComponent(ReconopenroDetails);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
