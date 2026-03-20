import { Routes } from '@angular/router';

export const routes: Routes = [
  { path: 'Header', loadComponent: () => import('./Layout/header/header').then(m => m.Header) },
  // ACCOUNTING BLOCK
  { path: 'AccountMapping', loadComponent: () => import('./Reports/Accounting/AccountMapping/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'FinancialStatement', loadComponent: () => import('./Reports/Accounting/FinancialStatement/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'FinancialSummary', loadComponent: () => import('./Reports/Accounting/FinancialSummary/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'ExpenseTrend', loadComponent: () => import('./Reports/Accounting/ExpenseTrend/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'FinancialStatementDetails', loadComponent: () => import('./Reports/Accounting/FinancialStatementDetails/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'IncomeStatementStore', loadComponent: () => import('./Reports/Accounting/IncomeStatementStoreComposite/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'IncomeStatementTrend', loadComponent: () => import('./Reports/Accounting/IncomeStatementTrend/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'LendersMapping', loadComponent: () => import('./Reports/Accounting/LenderMapping/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'FixedIncomeByExpense', loadComponent: () => import('./Reports/Accounting/EnterpriseIncomeByExpense/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'VariableIncomeByExpense', loadComponent: () => import('./Reports/Accounting/EnterpriseIncomeByExpense/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'EnterpriseTracking', loadComponent: () => import('./Reports/Accounting/EnterpriseTracking/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'LenderSummary', loadComponent: () => import('./Reports/Accounting/LenderSummary/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'TitleReport', loadComponent: () => import('./Reports/Accounting/TitleReport/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'LoanerInventoryV2', loadComponent: () => import('./Reports/Accounting/LoanerInventory/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'EnterpriseIncomeByExpense', loadComponent: () => import('./Reports/Accounting/EnterpriseIncomeByExpense/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'EnterpriseExpense', loadComponent: () => import('./Reports/Accounting/EnterpriseIncomeByExpenseDepartment/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'EnterpriseIncomeByExpenseTrend', loadComponent: () => import('./Reports/Accounting/EnterpriseIncomeByExpenseTrend/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'SellingGross', loadComponent: () => import('./Reports/Accounting/SellingGross/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'NetAddsDeducts', loadComponent: () => import('./Reports/Accounting/NetAddsByDeducts/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'PALenderSummary', loadComponent: () => import('./Reports/Accounting/PALenderReport/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'UsedCarWholesale', loadComponent: () => import('./Reports/Accounting/UsedCarWholesale/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'GLLookup', loadComponent: () => import('./Reports/Accounting/GLLookUp/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'WholesaleGrossGroup', loadComponent: () => import('./Reports/Accounting/WholesaleGrossGroup/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'RetailGrossGroup', loadComponent: () => import('./Reports/Accounting/RetailGrossGroup/dashboard/dashboard').then(m => m.Dashboard) },

  // RECEIVABLES BLOCK
  { path: 'Receivables', loadComponent: () => import('./Reports/Others/Receivables/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'BookedDeals', loadComponent: () => import('./Reports/Others/BookedDeals/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'CIT', loadComponent: () => import('./Reports/Others/CITFloorplan/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'WarrantyReceivables', loadComponent: () => import('./Reports/Others/WarrantyReceivables/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'VehicleReceivables', loadComponent: () => import('./Reports/Others/VehicleReceivables/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'FactoryIncentiveReceivables', loadComponent: () => import('./Reports/Others/FactoryIncentiveReceivables/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'Wholesale', loadComponent: () => import('./Reports/Others/Receivables/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'FinanceReserve', loadComponent: () => import('./Reports/Others/Receivables/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'Employee', loadComponent: () => import('./Reports/Others/Receivables/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'OpenAccountReceivables', loadComponent: () => import('./Reports/Others/OpenAccountsReceivables/dashboard/dashboard').then(m => m.Dashboard) },

  //LIABILITIES
  { path: 'Liabilities', loadComponent: () => import('./Reports/Others/Liabilities/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'NewFlooring', loadComponent: () => import('./Reports/Others/Liabilities/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'RentalInventory', loadComponent: () => import('./Reports/Others/Liabilities/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'TT&L', loadComponent: () => import('./Reports/Others/Liabilities/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'LienPayoffs', loadComponent: () => import('./Reports/Others/Liabilities/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'WeOwe', loadComponent: () => import('./Reports/Others/Liabilities/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'FinanceProductLiabilities', loadComponent: () => import('./Reports/Others/FinanceProductLiabilities/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'TitleTracking', loadComponent: () => import('./Reports/Others/TitleTracking/dashboard/dashboard').then(m => m.Dashboard) },


  // SALES BLOCK
  { path: 'SalesGross', loadComponent: () => import('./Reports/Sales/SalesGross/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'SalesGrossManager', loadComponent: () => import('./Reports/Sales/SalesGrossManager/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'SalesGrossExecutive', loadComponent: () => import('./Reports/Sales/SalesGrossExecutive/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'CarDeals', loadComponent: () => import('./Reports/Sales/CarDeals/dashboard/dashboard').then(m => m.Dashboard) },
  // { path: 'SalespersonRanking', loadComponent: () => import('./Reports/Sales/SalesPersonRanking/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'SPranking', loadComponent: () => import('./Reports/Sales/SalesPersonRanking/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'FandIManagerRanking', loadComponent: () => import('./Reports/Sales/FandIManagerRanking/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'InventoryBookV1', loadComponent: () => import('./Reports/Sales/InventoryBook/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'InventorySummaryV1', loadComponent: () => import('./Reports/Sales/InventorySummary/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'Appointments', loadComponent: () => import('./Reports/Sales/Appointments/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'VariableGrossGL', loadComponent: () => import('./Reports/Sales/VariableGrossGL/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'QuickInventory', loadComponent: () => import('./Reports/Sales/QuickInventoryReport/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'SalesReconciliation', loadComponent: () => import('./Reports/Sales/SalesReconciliation/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'BookedToFinalReconciliation', loadComponent: () => import('./Reports/Sales/BookedToClosedReconciliation/dashboard/dashboard').then(m => m.Dashboard) },

  { path: 'FandIProductPenetration', loadComponent: () => import('./Reports/Sales/FandIProductPenetration/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'FandIProductPenetrationV2', loadComponent: () => import('./Reports/Others/FandIProductPenetration/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'FandISummary', loadComponent: () => import('./Reports/Sales/FandISummary/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'SalesContest', loadComponent: () => import('./Reports/Sales/SalesContest/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'UsedVehicleStocking', loadComponent: () => import('./Reports/Sales/UsedVehicleStocking/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'TradeWinPercentage', loadComponent: () => import('./Reports/Sales/TradeWinPercentage/dashboard/dashboard').then(m => m.Dashboard) },


  { path: 'SalesTax', loadComponent: () => import('./Reports/Sales/SalesTax/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'UsedWriteDown', loadComponent: () => import('./Reports/Sales/InventoryValuation/dashboard/dashboard').then(m => m.Dashboard) },

  // SERVICE BLOCK
  { path: 'ServiceAdvisorRanking', loadComponent: () => import('./Reports/Services/ServiceAdvisorRanking/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'ServiceAppointments', loadComponent: () => import('./Reports/Services/ServiceAppointments/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'ServiceGross', loadComponent: () => import('./Reports/Services/ServiceGross/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'ServiceGrossGL', loadComponent: () => import('./Reports/Services/ServiceGrossGL/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'FixedSummaryGL', loadComponent: () => import('./Reports/Services/ServiceSummaryGL/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'ROTraffic', loadComponent: () => import('./Reports/Services/ROTraffic/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'TireGroupSalesReport', loadComponent: () => import('./Reports/Services/TireGroupSalesReport/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'ServiceOpenRO', loadComponent: () => import('./Reports/Services/ServiceOpenRO/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'InventoryOpenRO', loadComponent: () => import('./Reports/Services/InventoryOpenRO/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'ServiceAbsorption', loadComponent: () => import('./Reports/Services/ServiceAbsorption/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'ServiceLaborTrendDetail', loadComponent: () => import('./Reports/Services/ServiceLaborTrendDetails/dashboard/dashboard').then(m => m.Dashboard) },

  // OTHERS

  { path: 'FinanceReserveRecon', loadComponent: () => import('./Reports/Others/FinaceReserveReconReport/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'CollisionCenter', loadComponent: () => import('./Reports/Others/CollisionCenter/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'BudgetForecastInput', loadComponent: () => import('./Reports/Others/BudgetForecastInput/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'BudgetForecastInputAdd', loadComponent: () => import('./Reports/Others/BudgetForecastInput/budget-forecast-input-variables/budget-forecast-input-variables').then(m => m.BudgetForecastInputVariables) },
  { path: 'LeadSourceReport', loadComponent: () => import('./Reports/Others/LeadSourceReport/dashboard/dashboard').then(m => m.Dashboard) },

  //INVENTORY BLOCK
  { path: 'InventoryBrowser', loadComponent: () => import('./Reports/Inventory/InventoryBrowser/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'InventoryBook', loadComponent: () => import('./Reports/Inventory/InventoryBook/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'InventorySummary', loadComponent: () => import('./Reports/Inventory/InventorySummary/dashboard/dashboard').then(m => m.Dashboard) },

  { path: 'PartsAging', loadComponent: () => import('./Reports/Parts/PartsAging/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'PartsGrossGL', loadComponent: () => import('./Reports/Parts/PartsGrossGL/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'PartsSummaryGL', loadComponent: () => import('./Reports/Parts/PartsSummaryGL/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'PartsOpenRO', loadComponent: () => import('./Reports/Parts/PartsOpenRO/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'PartsGrossProfitPerformanceScoreCard', loadComponent: () => import('./Reports/Parts/PartsGrossProfitPerformanceScoreCard/dashboard/dashboard').then(m => m.Dashboard) },
  { path: 'PartsTrendingReport', loadComponent: () => import('./Reports/Parts/PartsTrendingReport/dashboard/dashboard').then(m => m.Dashboard) },

  { path: 'SalesManagerRanking', loadComponent: () => import('./Reports/Sales/SalesManangerRanking/dashboard/dashboard').then(m => m.Dashboard) },

];

