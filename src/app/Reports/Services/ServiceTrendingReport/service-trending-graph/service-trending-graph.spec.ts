import { ComponentFixture, TestBed } from '@angular/core/testing';

import { ServiceTrendingGraph } from './service-trending-graph';

describe('ServiceTrendingGraph', () => {
  let component: ServiceTrendingGraph;
  let fixture: ComponentFixture<ServiceTrendingGraph>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [ServiceTrendingGraph]
    })
    .compileComponents();

    fixture = TestBed.createComponent(ServiceTrendingGraph);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
