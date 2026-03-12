import { Component, OnInit, signal } from '@angular/core';
import { RouterOutlet } from '@angular/router';
import { Header } from './Layout/header/header';
import { Externalmenu } from './Layout/externalmenu/externalmenu'
import { CommonModule } from '@angular/common';
import { SidebarService } from '../app/Core/Providers/Shared/sidebar.service (1)';
import { ToastContainer } from './Layout/toast-container/toast-container';

@Component({
  selector: 'app-root',
  standalone: true,
  imports: [RouterOutlet, CommonModule, Header, Externalmenu, ToastContainer],
  templateUrl: './app.html',
  styleUrls: ['./app.scss']
})
export class App implements OnInit {
  isHeaderReady = false;
  isSidebarCollapsed = false;

  protected readonly title = signal('Fredbeans');

  constructor(private sidebarService: SidebarService) { }
  ngOnInit() {
    this.sidebarService.isCollapsed$.subscribe((collapsed) => {
      this.isSidebarCollapsed = collapsed;
    });
  }

  onHeaderReady(): void {
    this.isHeaderReady = true;
    console.log('Dashboard received Header ready event');
  }
}
