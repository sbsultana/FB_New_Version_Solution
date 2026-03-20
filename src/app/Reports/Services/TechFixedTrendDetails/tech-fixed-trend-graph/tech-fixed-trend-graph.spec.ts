import { ComponentFixture, TestBed } from '@angular/core/testing';

import { TechFixedTrendGraph } from './tech-fixed-trend-graph';

describe('TechFixedTrendGraph', () => {
  let component: TechFixedTrendGraph;
  let fixture: ComponentFixture<TechFixedTrendGraph>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [TechFixedTrendGraph]
    })
    .compileComponents();

    fixture = TestBed.createComponent(TechFixedTrendGraph);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
