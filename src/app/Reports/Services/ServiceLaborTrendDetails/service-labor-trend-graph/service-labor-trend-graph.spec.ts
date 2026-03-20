import { ComponentFixture, TestBed } from '@angular/core/testing';

import { ServiceLaborTrendGraph } from './service-labor-trend-graph';

describe('ServiceLaborTrendGraph', () => {
  let component: ServiceLaborTrendGraph;
  let fixture: ComponentFixture<ServiceLaborTrendGraph>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [ServiceLaborTrendGraph]
    })
    .compileComponents();

    fixture = TestBed.createComponent(ServiceLaborTrendGraph);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
