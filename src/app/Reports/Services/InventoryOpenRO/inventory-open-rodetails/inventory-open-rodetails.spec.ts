import { ComponentFixture, TestBed } from '@angular/core/testing';

import { InventoryOpenRODetails } from './inventory-open-rodetails';

describe('InventoryOpenRODetails', () => {
  let component: InventoryOpenRODetails;
  let fixture: ComponentFixture<InventoryOpenRODetails>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [InventoryOpenRODetails]
    })
    .compileComponents();

    fixture = TestBed.createComponent(InventoryOpenRODetails);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
