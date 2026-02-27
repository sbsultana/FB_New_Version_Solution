import { Routes } from '@angular/router';

export const routes: Routes = [
  { path: 'Header', loadComponent: () => import('./Layout/header/header').then(m => m.Header) },
  // ACCOUNTING BLOCK
  { path: 'AccountMapping', loadComponent: () => import('./Reports/Accounting/AccountMapping/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'FinancialStatement', loadComponent: () => import('./Reports/Accounting/FinancialStatement/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'financialsummary', loadComponent: () => import('./Reports/Accounting/FinancialSummary/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'ExpenseTrend', loadComponent: () => import('./Reports/Accounting/ExpenseTrend/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'IncomeStatementTrend', loadComponent: () => import('./Reports/Accounting/IncomeStatementTrend/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'IncomeStatementStoreComposite', loadComponent: () => import('./Reports/Accounting/IncomeStatement/income-statement-store-composite/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'IncomeStatementStore', loadComponent: () => import('./Reports/Accounting/IncomeStatementStoreComposite/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'IncomeStatementStoreV2', loadComponent: () => import('./Reports/Accounting/IncomeStatementStoreCompositeV2/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'EnterpriseTracking', loadComponent: () => import('./Reports/Accounting/EnterpriseTracking/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'EnterpriseIncomeByExpense', loadComponent: () => import('./Reports/Accounting/EnterpriseIncomeByExpense/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'FixedIncomeByExpense', loadComponent: () => import('./Reports/Accounting/EnterpriseIncomeByExpense/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'VariableIncomeByExpense', loadComponent: () => import('./Reports/Accounting/EnterpriseIncomeByExpense/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'EnterpriseIncomeByExpenseTrend', loadComponent: () => import('./Reports/Accounting/EnterpriseIncomeByExpenseTrend/dashboard/dashboard').then(m => m.Dashboard) },


  // SALES BLOCK
  { path: 'SalesGross', loadComponent: () => import('./Reports/Sales/SalesGross/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'CarDeals', loadComponent: () => import('./Reports/Sales/CarDeals/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'SalesPersonRanking', loadComponent: () => import('./Reports/Sales/SalesPersonRanking/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'FandIManagerRanking', loadComponent: () => import('./Reports/Sales/FandIManagerRanking/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'InventoryBook', loadComponent: () => import('./Reports/Sales/InventoryBook/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'InventorySummary', loadComponent: () => import('./Reports/Sales/InventorySummary/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'Appointments', loadComponent: () => import('./Reports/Sales/Appointments/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'VariableGrossGL', loadComponent: () => import('./Reports/Sales/VariableGrossGL/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'QuickInventory', loadComponent: () => import('./Reports/Sales/QuickInventoryReport/dashboard/dashboard').then(m => m.Dashboard) },

  // SERVICE BLOCK
  { path: 'ServiceAdvisorRanking', loadComponent: () => import('./Reports/Services/ServiceAdvisorRanking/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'ServiceAppointments', loadComponent: () => import('./Reports/Services/ServiceAppointments/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'ServiceGrossOld', loadComponent: () => import('./Reports/Services/ServiceGross/dashboard/dashboard').then(m => m.Dashboard) },

  // OTHERS
  { path: 'IncomeBudgetForecast', loadComponent: () => import('./Reports/Others/IncomeBudgetForecast/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'CITFloorplan', loadComponent: () => import('./Reports/Others/CITFloorplan/dashboard/dashboard').then(m => m.Dashboard) },

  // parts
  { path: 'PartsAging', loadComponent: () => import('./Reports/Parts/PartsAging/dashboard/dashboard').then(m => m.Dashboard) },

];
