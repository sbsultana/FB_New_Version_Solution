import { ComponentFixture, TestBed } from '@angular/core/testing';

import { PartsTrendingGraph } from './parts-trending-graph';

describe('PartsTrendingGraph', () => {
  let component: PartsTrendingGraph;
  let fixture: ComponentFixture<PartsTrendingGraph>;

  beforeEach(async () => {
    await TestBed.configureTestingModule({
      imports: [PartsTrendingGraph]
    })
    .compileComponents();

    fixture = TestBed.createComponent(PartsTrendingGraph);
    component = fixture.componentInstance;
    fixture.detectChanges();
  });

  it('should create', () => {
    expect(component).toBeTruthy();
  });
});
