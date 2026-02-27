import { NgFor } from '@angular/common';
import { Component, ElementRef, ViewChild } from '@angular/core';


@Component({
  selector: 'app-dashboard',
  standalone: true,
  imports: [NgFor],
  templateUrl: './dashboard.html',
  styleUrl: './dashboard.scss'
})
export class Dashboard {

  unitsData = [
    { type: 'New', mtd: 17, pace: 212, target: 303, diff: -91 },
    { type: 'Used', mtd: 4, pace: 85, target: 120, diff: -35 },
    { type: 'Total', mtd: 21, pace: 297, target: 423, diff: -126 },
    { type: 'New', mtd: 17, pace: 212, target: 303, diff: -91 },
    { type: 'Used', mtd: 4, pace: 85, target: 120, diff: -35 },
    { type: 'Total', mtd: 21, pace: 297, target: 423, diff: -126 },
    { type: 'New', mtd: 17, pace: 212, target: 303, diff: -91 },
    { type: 'Used', mtd: 4, pace: 85, target: 120, diff: -35 },
    { type: 'Total', mtd: 21, pace: 297, target: 423, diff: -126 },
    { type: 'New', mtd: 17, pace: 212, target: 303, diff: -91 },
    { type: 'Used', mtd: 4, pace: 85, target: 120, diff: -35 },
    { type: 'Total', mtd: 21, pace: 297, target: 423, diff: -126 },
    { type: 'New', mtd: 17, pace: 212, target: 303, diff: -91 },
    { type: 'Used', mtd: 4, pace: 85, target: 120, diff: -35 },
    { type: 'Total', mtd: 21, pace: 297, target: 423, diff: -126 },
    { type: 'New', mtd: 17, pace: 212, target: 303, diff: -91 },
    { type: 'Used', mtd: 4, pace: 85, target: 120, diff: -35 },
    { type: 'Total', mtd: 21, pace: 297, target: 423, diff: -126 },
    { type: 'New', mtd: 17, pace: 212, target: 303, diff: -91 },
    { type: 'Used', mtd: 4, pace: 85, target: 120, diff: -35 },
    { type: 'Total', mtd: 21, pace: 297, target: 423, diff: -126 },
    { type: 'New', mtd: 17, pace: 212, target: 303, diff: -91 },
    { type: 'Used', mtd: 4, pace: 85, target: 120, diff: -35 },
    { type: 'Total', mtd: 21, pace: 297, target: 423, diff: -126 },
    { type: 'New', mtd: 17, pace: 212, target: 303, diff: -91 },
    { type: 'Used', mtd: 4, pace: 85, target: 120, diff: -35 },
    { type: 'Total', mtd: 21, pace: 297, target: 423, diff: -126 },
    { type: 'New', mtd: 17, pace: 212, target: 303, diff: -91 },
    { type: 'Used', mtd: 4, pace: 85, target: 120, diff: -35 },
    { type: 'Total', mtd: 21, pace: 297, target: 423, diff: -126 },
    { type: 'New', mtd: 17, pace: 212, target: 303, diff: -91 },
    { type: 'Used', mtd: 4, pace: 85, target: 120, diff: -35 },
    { type: 'Total', mtd: 21, pace: 297, target: 423, diff: -126 },
    { type: 'New', mtd: 17, pace: 212, target: 303, diff: -91 },
    { type: 'Used', mtd: 4, pace: 85, target: 120, diff: -35 },
    { type: 'Total', mtd: 21, pace: 297, target: 423, diff: -126 },
    { type: 'New', mtd: 17, pace: 212, target: 303, diff: -91 },
    { type: 'Used', mtd: 4, pace: 85, target: 120, diff: -35 },
    { type: 'Total', mtd: 21, pace: 297, target: 423, diff: -126 },
    { type: 'New', mtd: 17, pace: 212, target: 303, diff: -91 },
    { type: 'Used', mtd: 4, pace: 85, target: 120, diff: -35 },
    { type: 'Total', mtd: 21, pace: 297, target: 423, diff: -126 },
    { type: 'New', mtd: 17, pace: 212, target: 303, diff: -91 },
    { type: 'Used', mtd: 4, pace: 85, target: 120, diff: -35 },
    { type: 'Total', mtd: 21, pace: 297, target: 423, diff: -126 },
    { type: 'New', mtd: 17, pace: 212, target: 303, diff: -91 },
    { type: 'Used', mtd: 4, pace: 85, target: 120, diff: -35 },
    { type: 'Total', mtd: 21, pace: 297, target: 423, diff: -126 },
  ];



  totalGross = { target: 250000, diff: -78358 };

  printSection() {
    window.print();
  }

}