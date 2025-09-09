import { AfterViewInit, Component, Injector, OnInit, ViewChild, effect, runInInjectionContext, signal, afterNextRender } from '@angular/core';
import { BudgetPlanService } from '../../services/budget-plan-service';
import { BudgetPlanUpsert } from '../../models/BudgetPlanModel';
import { KENDO_SPREADSHEET, SheetDescriptor, SpreadsheetComponent } from '@progress/kendo-angular-spreadsheet';
import { firstValueFrom } from 'rxjs';

@Component({
  selector: 'app-spreadsheet',
  standalone: true,
  imports: [KENDO_SPREADSHEET],
  templateUrl: './spreadsheet.html',
  styleUrl: './spreadsheet.scss'
})
export class SpreadsheetWrapperComponent implements OnInit, AfterViewInit {

  @ViewChild('spreadsheet') spreadsheet!: SpreadsheetComponent;
  sheets = signal<SheetDescriptor[]>([]);
  dropdowns = signal<any>({});
  dataLoaded = signal(false);
  private appliedDropdowns = false;
  private injector: Injector;
  private lastFormattedRowCount = 0; // Track formatting state

  // Lookup maps
  private projectCodeToName: Record<string,string> = {};
  private projectNameToCode: Record<string,string> = {};
  private employeeCodeToName: Record<string,string> = {};
  private employeeNameToCode: Record<string,string> = {};
  private monthIdToName: Record<number,string> = {};
  private monthNameToId: Record<string,number> = {};
  private statusNameToValue: Record<string,string> = {};

  constructor(private service: BudgetPlanService, injector: Injector) {
    this.injector = injector;

    // Reactively apply dropdowns once spreadsheet and dropdown data are ready
    effect(() => {
      const ss = this.spreadsheet;
      const dl = this.dataLoaded();
      const dd = this.dropdowns();
      if (ss && dl && Object.keys(dd).length > 0) {
        this.appliedDropdowns = true; 
        runInInjectionContext(this.injector, () => {
          afterNextRender(() => {
            this.applyDropdowns();
            setTimeout(() => this.attachFormulasAndFormatting(), 100);
          });
        });
      }
    });
  }

  async ngOnInit() {
    try {
      const [projects, employees, statuses, months] = await Promise.all([
        firstValueFrom(this.service.getProjects()),
        firstValueFrom(this.service.getEmployees()),
        firstValueFrom(this.service.getStatuses()),
        firstValueFrom(this.service.getMonths())
      ]);

      // Build lookup maps
      projects.forEach(p => {
        const display = p.name ?? p.code;
        this.projectCodeToName[p.code] = display;
        this.projectNameToCode[display] = p.code;
      });

      employees.forEach(e => {
        const display = e.name ?? e.employeeCode;
        this.employeeCodeToName[e.employeeCode] = display;
        this.employeeNameToCode[display] = e.employeeCode;
      });

      months.forEach(m => {
        const display = m.name ?? String(m.monthId);
        this.monthIdToName[m.monthId] = display;
        this.monthNameToId[display] = m.monthId;
      });

      statuses.forEach(s => {
        const display = s.name;
        this.statusNameToValue[display] = display;
      });

      this.dropdowns.set({ projects, employees, statuses, months });
      this.loadData();
    } catch (error) {
      console.error('Error loading dropdown data:', error);
    }
  }

  ngAfterViewInit() {}

  private rowToId: Record<number, number> = {}; // row index -> id

  loadData() {
    this.service.getPlans({ page: 1, pageSize: 100 }).subscribe(res => {
    this.rowToId = {}; // reset map

    const rows = res.items.map((p, idx) => {
      const rowIndex = idx + 2; // +2 because header is row 1
      this.rowToId[rowIndex] = p.budgetPlanId;

      return {
        cells: [
          { value: this.projectCodeToName[p.projectCode] ?? p.projectCode },
          { value: this.employeeCodeToName[p.employeeCode] ?? p.employeeCode },
          { value: p.year },
          { value: this.monthIdToName[p.monthId] ?? p.monthId },
          { value: p.budgetAllocated },
          { value: p.hoursPlanned },
          { value: p.cost ?? null },
          { value: this.statusNameToValue[p.status] ?? p.status },
          { value: p.comments ?? '' }
        ]
      };
      });

      const header = { cells: [
        { value: 'Project', bold: true },
        { value: 'Employee', bold: true },
        { value: 'Year', bold: true },
        { value: 'Month', bold: true },
        { value: 'Budget', bold: true },
        { value: 'Hours', bold: true },
        { value: 'Cost', bold: true },
        { value: 'Status', bold: true },
        { value: 'Comments', bold: true }
      ] };

      const sheet: SheetDescriptor = {
        name: 'Budget Plans',
        rows: [ header, ...rows ],
        columns: [
          { width: 160 }, { width: 160 }, { width: 80 }, { width: 100 },
          { width: 120 }, { width: 120 }, { width: 120 }, { width: 120 }, { width: 250 }
        ]
      };

      this.sheets.set([sheet]);
      this.dataLoaded.set(true);
      this.lastFormattedRowCount = 0; // Reset formatting state
    });
  }

  private attachFormulasAndFormatting() {
    const widget: any = (this.spreadsheet as any).spreadsheetWidget;
    if (!widget) return;

    const sheet: any = widget.activeSheet();
    if (!sheet) return;

    // console.log('Attaching formulas and formatting');

    // Apply initial formatting
    this.applyConditionalFormatting(sheet);

    // Bind change event to apply conditional formatting dynamically
    // Remove existing handlers first to prevent multiple bindings
    widget.unbind && widget.unbind('change');
    widget.bind && widget.bind('change', () => {
      const sheet = widget.activeSheet();
      if (sheet) {
        // Debounce formatting to prevent excessive calls
        setTimeout(() => this.applyConditionalFormatting(sheet), 50);
      }
    });
  }

  private getStatusColor(status: string): string {
    const s = status?.toLowerCase().trim();
    if (!s) return '';
    
    // Match exact backend status values (case-insensitive)
    if (s === 'approved') return '#28a745';     // Green
    if (s === 'planned') return '#ffc107';     // Yellow/Amber
    if (s === 'over budget') return '#dc3545';  // Red
    
    return '';
  }

  private getLightStatusColor(status: string): string {
    const s = status?.toLowerCase().trim();
    if (!s) return '';
    
    if (s === 'approved') return '#d4edda';     // Light green
    if (s === 'planned') return '#fff3cd';     // Light yellow
    if (s === 'over budget') return '#f8d7da';  // Light red
    
    return '';
  }

  private findLastDataRow(sheet: any): number {
    try {
      // More efficient approach: check from a reasonable starting point
      let lastRow = 1; // Start from header row
      const maxCheck = 200; // Reasonable limit instead of 1000
      
      for (let r = 2; r <= maxCheck; r++) {
        try {
          // Check if any cell in the row has meaningful content
          const hasContent = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'].some(col => {
            try {
              const cellValue = sheet.range(`${col}${r}`).value();
              return cellValue !== null && cellValue !== undefined && cellValue !== '';
            } catch {
              return false;
            }
          });
          
          if (hasContent) {
            lastRow = r;
          }
        } catch (e) {
          // If we hit an error, stop checking further
          break;
        }
      }
      
      // console.log(`Found last data row: ${lastRow}`);
      return lastRow;
    } catch (error) {
      console.error('Error finding last data row:', error);
      return 1;
    }
  }

  private applyConditionalFormatting(sheet: any) {
  if (!sheet) return;
  
  // console.log('Applying conditional formatting');
  
  try {
    const lastRow = this.findLastDataRow(sheet);
    const firstDataRow = 2;
    
    if (lastRow < firstDataRow) {
      // console.log('No data rows to format');
      return;
    }

    // Only format new rows or if we haven't formatted before
    // const startRow = Math.max(firstDataRow, this.lastFormattedRowCount + 1);
    const startRow = firstDataRow; // Always start from first data row

    
    // console.log(`Formatting rows ${startRow} to ${lastRow}`);

    // Apply formulas and number formatting ONLY to new rows
    for (let r = startRow; r <= lastRow; r++) {
      try {
        // Apply formula for cost calculation (G = E * F)
        const budgetCell = sheet.range(`E${r}`);
        const hoursCell = sheet.range(`F${r}`);
        const costCell = sheet.range(`G${r}`);
        
        if (budgetCell.value() !== null && hoursCell.value() !== null) {
          costCell.formula(`=E${r}*F${r}`);
        }
        
        // Apply number formatting
        sheet.range(`E${r}:G${r}`).format('0.00');
      } catch (formatError) {
        console.warn(`Error applying formula/format to row ${r}:`, formatError);
      }
    }

    // Apply conditional formatting to ALL rows (to handle changes in existing rows)
    for (let r = firstDataRow; r <= lastRow; r++) {
      try {
        this.formatRow(sheet, r);
      } catch (error) {
        console.error(`Error applying formatting to row ${r}:`, error);
      }
    }
    
    this.lastFormattedRowCount = lastRow;
    
  } catch (error) {
    console.error('Error in applyConditionalFormatting:', error);
  }
}

private getSelectedRowIndices(): number[] {
  const widget: any = (this.spreadsheet as any).spreadsheetWidget;
  if (!widget) return [];

  const sheet = widget.activeSheet();
  if (!sheet) return [];

  const selection = sheet.selection();
  if (!selection) return [];

  const startRow = selection.topLeft().row + 1; // 0-based → 1-based
  const endRow = selection.bottomRight().row + 1;

  const rows: number[] = [];
  for (let r = startRow; r <= endRow; r++) {
    rows.push(r);
  }
  return rows;
}


  private formatRow(sheet: any, rowNum: number) {
    try {
      const statusCell = sheet.range(`H${rowNum}`);
      const rawValue = statusCell?.value();
      const status = (rawValue ?? '').toString().trim();

      // Don't reset formatting too aggressively - only reset what we need to change
      const rowRange = sheet.range(`A${rowNum}:I${rowNum}`);
      
      // Apply row coloring based on status
      const colorHex = this.getStatusColor(status);
      
      if (colorHex) {
        const lightColor = this.getLightStatusColor(status);
        // Apply light background to entire row
        rowRange.background(lightColor);
        // Apply dark background with white text to status cell
        statusCell.background(colorHex).color('#ffffff').bold(true);
      } else {
        // Reset to default if no status color
        rowRange.background(null);
        statusCell.color(null).bold(null);
      }

      // Highlight cost if over budget
      const budget = Number(sheet.range(`E${rowNum}`).value() ?? 0);
      const cost = Number(sheet.range(`G${rowNum}`).value() ?? 0);
      const costCell = sheet.range(`G${rowNum}`);
      
      if (!isNaN(cost) && !isNaN(budget) && cost > budget && budget > 0) {
        costCell.color('#721c24').bold(true).background('#f8d7da');
      } else if (!colorHex) {
        // Only reset if we don't have a status color
        costCell.color(null).bold(null).background(null);
      }
      
    } catch (error) {
      console.error(`Error formatting row ${rowNum}:`, error);
    }
  }

  applyDropdowns() {
    // console.log('Applying dropdowns...');
    const spreadsheet = this.spreadsheet;
    if (!spreadsheet) {
      console.error('Spreadsheet not available');
      return;
    }
    
    const widget: any = (spreadsheet as any).spreadsheetWidget;
    if (!widget) {
      console.error('Spreadsheet widget not available');
      return;
    }
    
    const sheet: any = widget.activeSheet();
    if (!sheet) {
      console.error('Active sheet not available');
      return;
    }

    try {
      const dropdownData = this.dropdowns();
      if (!dropdownData || Object.keys(dropdownData).length === 0) {
        console.error('Dropdown data not loaded');
        return;
      }

      const projectNames = (dropdownData.projects || []).map((p: { code: string; name: string }) => p.name ?? p.code);
      const employeeNames = (dropdownData.employees || []).map((e: { employeeCode: string; name: string }) => e.name ?? e.employeeCode);
      const monthNames = (dropdownData.months || []).map((m: { monthId: number; name: string }) => m.name ?? String(m.monthId));
      const statusNames = (dropdownData.statuses || []).map((s: { name: string }) => s.name);

      // console.log('Dropdown data:', { projectNames: projectNames.length, employeeNames: employeeNames.length, monthNames: monthNames.length, statusNames: statusNames.length });

      const setList = (range: string, values: string[]) => {
        if (!values.length) {
          console.warn(`No values for range ${range}`);
          return;
        }
        try {
          sheet.range(range).validation({
            dataType: 'list',
            from: `"${values.join(',')}"`,
            allowNulls: false,
            showButton: true,
            comparerType: 'list',
            type: 'reject'
          });
          // console.log(`Applied list validation to ${range} with ${values.length} items`);
        } catch (error) {
          console.error(`Error setting list for ${range}:`, error);
        }
      };

      const setCustom = (range: string, formula: string, allowNulls: boolean = true) => {
        try {
          sheet.range(range).validation({
            dataType: 'custom',
            comparerType: 'custom',
            from: formula,
            allowNulls,
            type: 'reject'
          });
          // console.log(`Applied custom validation to ${range}`);
        } catch (error) {
          console.error(`Error setting custom validation for ${range}:`, error);
        }
      };

      // Apply dropdown validations with extended ranges for new rows
      setList('A2:A', projectNames);    // Extended range
      setList('B2:B', employeeNames);   // Extended range
      setList('D2:D', monthNames);      // Extended range
      setList('H2:H', statusNames);     // Extended range

      setCustom('C2:C', 'AND(ISNUMBER(C2), C2>=2020, C2<=2030, INT(C2)=C2)', false);
      setCustom('E2:E', 'AND(ISNUMBER(E2), E2>0)', false);
      setCustom('F2:F', 'AND(ISNUMBER(F2), F2>=0, F2<=999)', false);
      setCustom('G2:G', 'OR(ISBLANK(G2), AND(ISNUMBER(G2), G2>0))', true);
      setCustom('I2:I', 'LEN(I2)<=500', true);

      // console.log('Dropdowns applied successfully');
      
    } catch (error) {
      console.error('Error applying dropdowns and validations:', error);
    }
  }

  saveData() {
    const spreadsheet = this.spreadsheet;
    if (!spreadsheet) return;

    const workbook = (spreadsheet as any).spreadsheetWidget.toJSON();
    const rows = workbook.sheets?.[0]?.rows ?? [];
    const data = rows.slice(1);

    const upserts: BudgetPlanUpsert[] = [];
    data.forEach((r : any) => {
      const c = r.cells ?? [];
      const projectDisplay = c[0]?.value;
      const employeeDisplay = c[1]?.value;

      if (projectDisplay && employeeDisplay) {
        const projectCode = this.projectNameToCode[projectDisplay] ?? (() => {
          const m = /.*\((.*)\)\s*$/.exec(projectDisplay);
          return m ? m[1] : projectDisplay;
        })();
        const employeeCode = this.employeeNameToCode[employeeDisplay] ?? employeeDisplay;

        upserts.push({
          projectCode,
          employeeCode,
          year: Number(c[2]?.value),
          monthId: Number(this.monthNameToId[c[3]?.value] ?? c[3]?.value),
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
        // Reset formatting state and reapply
        this.lastFormattedRowCount = 0;
        setTimeout(() => {
          this.attachFormulasAndFormatting();
        }, 100);
      }, err => {
        console.error('Save error', err);
        alert('Save failed. See console.');
      });
    } else {
      alert('No rows to save');
    }
  }


deleteSelected() {
  const selectedRows = this.getSelectedRowIndices();
  const idsToDelete = selectedRows
    .map(r => this.rowToId[r])
    .filter(id => !!id); // only valid ids

  if (!idsToDelete.length) {
    alert('No valid rows selected to delete.');
    return;
  }

  if (!confirm(`Delete ${idsToDelete.length} plan(s)?`)) return;

  this.service.bulkDelete(idsToDelete).subscribe({
    next: () => {
      alert('Deleted successfully!');
      this.loadData(); // reload grid
    },
    error: err => {
      console.error('Delete failed', err);
      alert('Delete failed. See console.');
    }
  });
}



}