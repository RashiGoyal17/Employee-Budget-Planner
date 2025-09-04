import { Component, OnInit, signal, viewChild } from '@angular/core';
import { BudgetPlanService } from '../../services/budget-plan-service';
import { BudgetPlanUpsert } from '../../models/BudgetPlanModel';
import { KENDO_SPREADSHEET, SheetDescriptor, SpreadsheetComponent } from '@progress/kendo-angular-spreadsheet';

@Component({
  selector: 'app-spreadsheet',
  standalone:true,
  imports: [KENDO_SPREADSHEET],
  templateUrl: './spreadsheet.html',
  styleUrl: './spreadsheet.scss'
})
export class SpreadsheetWrapperComponent implements OnInit {


  spreadsheet = viewChild<SpreadsheetComponent>('spreadsheet');
  sheets = signal<SheetDescriptor[]>([]);
  dropdowns: any = {};   

  constructor(private service: BudgetPlanService) {}

  async ngOnInit() {
      const [projects, employees, statuses, months] = await Promise.all([
    this.service.getProjects().toPromise(),
    this.service.getEmployees().toPromise(),
    this.service.getStatuses().toPromise(),
    this.service.getMonths().toPromise()
  ]);
  this.dropdowns = { projects, employees, statuses, months };


    this.loadData();
  }

  loadData() {
    this.service.getPlans({ page: 1, pageSize: 50 }).subscribe(res => {
      const rows = res.items.map(p => ({
        cells: [
          { value: p.projectCode },
          { value: p.employeeCode },
          { value: p.year },
          { value: p.monthId },
          { value: p.budgetAllocated },
          { value: p.hoursPlanned },
          { value: p.cost },
          { value: p.status },
          { value: p.comments ?? '' }
        ]
      }));

      const sheet: SheetDescriptor = {
        name: 'Budget Plans',
        rows: [
          { cells: [
            { value: 'Project', bold: true },
            { value: 'Employee', bold: true },
            { value: 'Year', bold: true },
            { value: 'Month', bold: true },
            { value: 'Budget', bold: true },
            { value: 'Hours', bold: true },
            { value: 'Cost', bold: true },
            { value: 'Status', bold: true },
            { value: 'Comments', bold: true }
          ]},
          ...rows
        ],
                columns: [
          { width: 120 }, // Project
          { width: 120 }, // Employee
          { width: 80 },  // Year
          { width: 80 },  // Month
          { width: 100 }, // Budget
          { width: 100 }, // Hours
          { width: 100 }, // Cost
          { width: 100 }, // Status
          { width: 200 }  // Comments
        ],
      };

      this.sheets.set([sheet]);
      // Apply dropdown validation after sheet is created
      setTimeout(() => this.applyDropdowns(), 100);
    });
  }


applyDropdowns() {
  const spreadsheet = this.spreadsheet();
  if (!spreadsheet) return;

  const widget = spreadsheet.spreadsheetWidget;
  const sheet = widget.activeSheet();
  if (!sheet) return;

  // Cast ranges to any to access validation method
  (sheet.range("A2:A200") as any).validation({
    dataType: "list",
    from: `"${this.dropdowns.projects.map((p: any) => p.code).join(',')}"`,
    allowNulls: true
  });

  (sheet.range("B2:B200") as any).validation({
    dataType: "list",
    from: `"${this.dropdowns.employees.map((e: any) => e.employeeCode).join(',')}"`,
    allowNulls: true
  });

  (sheet.range("D2:D200") as any).validation({
    dataType: "list",
    from: `"${this.dropdowns.months.map((m: any) => m.monthId).join(',')}"`,
    allowNulls: true
  });

  (sheet.range("H2:H200") as any).validation({
    dataType: "list",
    from: `"${this.dropdowns.statuses.map((s: any) => s.name).join(',')}"`,
    allowNulls: true
  });
}


  saveData() {
    const spreadsheet = this.spreadsheet();
    if (!spreadsheet) return;

    const workbook = spreadsheet.spreadsheetWidget.toJSON();
    const rows = workbook.sheets?.[0]?.rows ?? [];
    const data = rows.slice(1);

    const upserts: BudgetPlanUpsert[] = [];
    data.forEach((r: any) => {
      const c = r.cells ?? [];
      if (c.length >= 8 && c[0]?.value && c[1]?.value) {
        upserts.push({
          projectCode: c[0]?.value,
          employeeCode: c[1]?.value,
          year: Number(c[2]?.value),
          monthId: Number(c[3]?.value),
          budgetAllocated: Number(c[4]?.value),
          hoursPlanned: Number(c[5]?.value),
          statusName: c[7]?.value,
          comments: c[8]?.value ?? ''
        });
      }
    });

    if (upserts.length) {
      this.service.bulkUpsert(upserts).subscribe(() => {
        alert('Saved successfully!');
        this.loadData(); // reload grid with fresh DB data
      });
    }
  }


}
