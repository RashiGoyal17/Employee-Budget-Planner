import { AfterViewInit, Component, Injector, OnInit, ViewChild, effect, runInInjectionContext, signal, afterNextRender, NgZone } from '@angular/core';
import { BudgetPlanService } from '../../services/budget-plan-service';
import { BudgetPlan, BudgetPlanUpsert } from '../../models/BudgetPlanModel';
import { KENDO_SPREADSHEET, SheetDescriptor, SpreadsheetComponent,SheetRow } from '@progress/kendo-angular-spreadsheet';
import { firstValueFrom } from 'rxjs';
import jsPDF from 'jspdf';
import autoTable from 'jspdf-autotable';
import * as XLSX from 'xlsx';
import { FormsModule } from '@angular/forms';

@Component({
  selector: 'app-spreadsheet',
  standalone: true,
  imports: [KENDO_SPREADSHEET,FormsModule],
  templateUrl: './spreadsheet.html',
  styleUrl: './spreadsheet.scss'
})
export class SpreadsheetWrapperComponent implements OnInit, AfterViewInit {

  @ViewChild('spreadsheet') spreadsheet!: SpreadsheetComponent;
  columnHeaders = signal<string[]>(['Project', 'Employee', 'Year', 'Month', 'Budget', 'Hours', 'Cost', 'Status', 'Comments']); // Custom headers
  sheets = signal<SheetDescriptor[]>([]);
  dropdowns = signal<any>({});
  dataLoaded = signal(false);
    // In your component class
  selectedRows = signal<number[]>([]);  // Track selected row indices reactively
  totalRecords = signal(0); // New: Track total
  private rowToId: Record<number, number> = {}; // row index -> id
  private appliedDropdowns = false;
  private injector: Injector;
  private lastFormattedRowCount = 0; // Track formatting state
  private currentData: BudgetPlan[] = []; // For local sort/filter


  // Filter/Sort state
  filterProject = '';
  filterEmployee = '';
  filterYear?: number;
  filterMonth?: string | number;  // Changed: allow string (name) or number (ID)
  sortBy = '';
  sortDir = 'asc';

  // Lookup maps
  private projectCodeToName: Record<string,string> = {};
  private projectNameToCode: Record<string,string> = {};
  private employeeCodeToName: Record<string,string> = {};
  private employeeNameToCode: Record<string,string> = {};
  private monthIdToName: Record<number,string> = {};
  private monthNameToId: Record<string,number> = {};
  private statusNameToValue: Record<string,string> = {};
  

  constructor(private service: BudgetPlanService, injector: Injector,private ngZone: NgZone) {
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
      await this.loadData(); // Await to ensure
    } catch (error) {
      console.error('Error loading dropdown data:', error);
    }
  }

  ngAfterViewInit() {
  // Ensure load if not done
    if (!this.dataLoaded()) {
      this.loadData();
    }
}


private freezeHeader() {
  const widget: any = (this.spreadsheet as any).spreadsheetWidget;
  if (widget) {
    const sheet: any = widget.activeSheet();
    if (sheet) {
      sheet.frozenRows(1); // Freeze the first row (header)
      console.log('Header frozen');
    }
  }
}


onSelectionChange(event: any): void {
  try {
    console.log('onSelectionChange fired with event:', event);

    const range = event?.range;
    if (!range || !range._ref || !range._ref.topLeft || !range._ref.bottomRight) {
      console.log('No valid range in event, clearing selection');
      this.selectedRows.set([]);
      return;
    }

    const topLeft = range._ref.topLeft;
    const bottomRight = range._ref.bottomRight;
    const startRow = topLeft.row + 1;  // 1-based (Kendo is 0-based)
    const endRow = bottomRight.row + 1;
    const rows: number[] = [];

    for (let r = startRow; r <= endRow; r++) {
      if (r >= 2) rows.push(r);  // skip header
    }

    this.selectedRows.set(rows);
    console.log('Selection updated:', rows);
  } catch (err) {
    console.error('Error in onSelectionChange:', err);
    this.selectedRows.set([]);
  }
}

async loadData() {
try {
    // Map month name to ID if filterMonth is string
    let monthIdNum: number | undefined;
    if (this.filterMonth && typeof this.filterMonth === 'string') {
      monthIdNum = this.monthNameToId[this.filterMonth as string];
    } else {
      monthIdNum = this.filterMonth as number;
    }

    const params: any = {
      ...(this.filterProject && { projectCodeOrName: this.filterProject }), // Updated param name
      ...(this.filterEmployee && { employeeCodeOrName: this.filterEmployee }), // Updated param name
      ...(this.filterYear && { year: this.filterYear }),
...(this.filterMonth !== undefined && (() => {
          if (typeof this.filterMonth === 'number') {
            return { monthId: this.filterMonth };
          } else {
            return { monthName: this.filterMonth as string };
          }
        })()), // Send numeric only
      ...(this.sortBy && { sortBy: this.sortBy }), // Now include sortBy
      ...(this.sortDir && { sortDir: this.sortDir })
    };
    console.log('API Params:', params); // Debug

    const res = await firstValueFrom(this.service.getPlans(params));
    console.log('Loaded data:', res);

      this.currentData = res.items; // For local ops
      this.totalRecords.set(res.total);
      this.rowToId = {};

      const rows = res.items.map((p, idx) => {
        const rowIndex = idx + 2;
        this.rowToId[rowIndex] = p.budgetPlanId;
        return {
          cells: [
            { value: this.projectCodeToName[p.projectCode] ?? p.projectCode },
            { value: this.employeeCodeToName[p.employeeCode] ?? p.employeeCode },
            { value: p.year },
            { value: this.monthIdToName[p.monthId] ?? p.monthId },
            { value: p.budgetAllocated },
            { value: p.hoursPlanned},
            { value: p.cost ?? null },
            { value: this.statusNameToValue[p.status] ?? p.status },
            { value: p.comments ?? '' }
          ]
        };
      });

      const headers = this.columnHeaders(); // Use custom
      const headerRow = { 
        cells: headers.map(h => ({ value: h, bold: true })) 
      };

      const sheet: SheetDescriptor = {
        name: 'Budget Plans',
        rows: [ headerRow, ...rows ],
        columns: [
          { width: 160 }, { width: 160 }, { width: 80 }, { width: 100 },
          { width: 120 }, { width: 120 }, { width: 120 }, { width: 120 }, { width: 250 }
        ]
      };

      this.sheets.set([sheet]);
      this.dataLoaded.set(true);
      this.lastFormattedRowCount = 0;
      // NEW: Apply dropdowns and freeze after a short delay to ensure sheet renders
  setTimeout(() => {
    this.applyDropdowns();
    this.attachFormulasAndFormatting();
  }, 150);
    } catch (error : any) {
    console.error('Load error:', error);
    if (error.status === 400) {
      console.error('Validation errors:', error.error?.errors); // Log details
    }
    alert(`Failed to load data: ${error.message || 'Check console for details.'}`);
  }
}


  // New: Apply Filter/Sort (Server-Side)
  applyFilterSort() {
    this.loadData();
  }

  // New: Clear Filter/Sort
  clearFilterSort() {
    this.filterProject = '';
    this.filterEmployee = '';
    this.filterYear = undefined;
    this.filterMonth = undefined;
    this.sortBy = '';
    this.sortDir = 'asc';
    this.loadData();
  }


// New helper: Case-insensitive comparison
private caseInsensitiveCompare(a: string, b: string): number {
  const lowerA = a.toLowerCase();
  const lowerB = b.toLowerCase();
  if (lowerA > lowerB) return 1;
  if (lowerA < lowerB) return -1;
  return 0; // Equal, preserve original order for stability
}

// Updated localSort (use display names, explicit direction)
localSort(field: string, direction: 'asc' | 'desc' = 'asc') { // Default to desc
  if (!this.currentData.length) return;

  this.currentData.sort((a, b) => {
    let valA: any = this.getFieldValue(a, field);
    let valB: any = this.getFieldValue(b, field);

    // Case-insensitive for strings
    if (typeof valA === 'string' && typeof valB === 'string') {
      const lowerA = valA.toLowerCase();
      const lowerB = valB.toLowerCase();
      return direction === 'desc'
        ? lowerB.localeCompare(lowerA) // Desc: Z->A
        : lowerA.localeCompare(lowerB); // Asc: A->Z
    }

    // Numbers: direct comparison
    const numCompare = (valA as number) - (valB as number);
    return direction === 'desc' ? -numCompare : numCompare;
  });

  this.updateSheetFromData();
}

// Update getFieldValue to return display names for sorting
private getFieldValue(item: BudgetPlan, field: string): any {
  switch (field) {
    case 'projectCode':
    case 'projectName': // Added alias for name-based sorting
      return this.projectCodeToName[item.projectCode] ?? item.projectCode;
    case 'employeeCode':
    case 'employeeName': // Added alias
      return this.employeeCodeToName[item.employeeCode] ?? item.employeeCode;
    case 'year': return item.year;
    case 'monthId': return item.monthId;
    case 'budgetAllocated': return item.budgetAllocated;
    case 'hoursPlanned': return item.hoursPlanned;
    case 'status': return item.status;
    default: return '';
  }
}

private updateSheetFromData() {
  // Map currentData to spreadsheet rows, reusing loadData's cell mapping
  const rows = this.currentData.map((p, idx) => {
    const rowIndex = idx + 2; // +2 because header is row 1
    this.rowToId[rowIndex] = p.budgetPlanId; // Update row-to-ID mapping

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

  // Use custom headers
  const headers = this.columnHeaders();
  const headerRow = { cells: headers.map(h => ({ value: h, bold: true })) };

  // Build sheet with same structure as loadData
  const sheet: SheetDescriptor = {
    name: 'Budget Plans',
    rows: [headerRow, ...rows],
    columns: [
      { width: 160 }, // Project
      { width: 160 }, // Employee
      { width: 80 },  // Year
      { width: 100 }, // Month
      { width: 120 }, // Budget
      { width: 120 }, // Hours
      { width: 120 }, // Cost
      { width: 120 }, // Status
      { width: 250 }  // Comments
    ]
  };

  // Update sheet signal
  this.sheets.set([sheet]);

  // Reapply formatting and validations
  setTimeout(() => this.attachFormulasAndFormatting(), 100);
  this.freezeHeader();
}

  // Update

  private attachFormulasAndFormatting() {
    const widget: any = (this.spreadsheet as any).spreadsheetWidget;
    if (!widget) return;

    const sheet: any = widget.activeSheet();
    if (!sheet) return;

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
      return lastRow;
    } catch (error) {
      console.error('Error finding last data row:', error);
      return 1;
    }
  }

  private applyConditionalFormatting(sheet: any) {
  if (!sheet) return;
  
  try {
    const lastRow = this.findLastDataRow(sheet);
    const firstDataRow = 2;
    
    if (lastRow < firstDataRow) {
      return;
    }

    // Only format new rows or if we haven't formatted before
    // const startRow = Math.max(firstDataRow, this.lastFormattedRowCount + 1);
    const startRow = firstDataRow; // Always start from first data row

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
  if (!widget) {
    console.log('No widget available');
    return [];
  }

  const sheet = widget.activeSheet();
  if (!sheet) {
    console.log('No active sheet');
    return [];
  }

  const selection = sheet.selection();
  if (!selection) {
    console.log('No selection');
    return [];
  }

  const startRow = selection.topLeft().row + 1;
  const endRow = selection.bottomRight().row + 1;
  console.log('Selection range:', { startRow, endRow });

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

// New: Edit Custom Headers
  editHeaders() {
    const newHeaders = prompt('Enter comma-separated headers (e.g., Project,Employee,Year,...):', this.columnHeaders().join(','));
    if (newHeaders) {
      this.columnHeaders.set(newHeaders.split(',').map(h => h.trim()));
      this.loadData(); // Reload to apply
    }
  }


  applyDropdowns() {
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

      setCustom('C2:C', 'AND(ISNUMBER(C2), C2>=1000, C2<=4050, INT(C2)=C2)', false);
      setCustom('E2:E', 'AND(ISNUMBER(E2), E2>0)', false);
      setCustom('F2:F', 'AND(ISNUMBER(F2), F2>=0, F2<=999)', false);
      setCustom('G2:G', 'OR(ISBLANK(G2), AND(ISNUMBER(G2), G2>0))', true);
      setCustom('I2:I', 'LEN(I2)<=500', true);
      this.freezeHeader();
      const widget: any = (this.spreadsheet as any).spreadsheetWidget;
      if (widget) {
        const sheet: any = widget.activeSheet();
        if (sheet) {
          // sheet.frozenRows(1); // Freeze header row
        }
      }
      this.freezeHeader();
    } catch (error) {
      console.error('Error applying dropdowns and validations:', error);
    }
  } 
    // } catch (error) {
    //   console.error('Error applying dropdowns and validations:', error);
    // }
  // }


saveData() {
  const spreadsheet = this.spreadsheet;
  if (!spreadsheet) return;

  const workbook = (spreadsheet as any).spreadsheetWidget.toJSON();
  const rows = workbook.sheets?.[0]?.rows ?? [];
  const data = rows.slice(1);

  const upserts: BudgetPlanUpsert[] = [];
  const invalidCells: { rowIndex: number, colIndex: number, reason: string }[] = [];
  const invalidRows: number[] = []; // Track invalid rows for feedback

  data.forEach((r: any, dataRowIndex: number) => {
    const rowIndex = dataRowIndex + 2; // 1-based row number for sheet
    const c = r.cells ?? [];
    const projectDisplay = c[0]?.value;
    const employeeDisplay = c[1]?.value;
    const yearVal = c[2]?.value;
    const monthDisplay = c[3]?.value;
    const budgetVal = c[4]?.value;
    const hoursVal = c[5]?.value;
    const statusName = c[7]?.value;
    const comments = c[8]?.value ?? '';

    // Map values (same as before)
    const projectCode = projectDisplay ? this.projectNameToCode[projectDisplay] : null;
    const employeeCode = employeeDisplay ? this.employeeNameToCode[employeeDisplay] : null;
    const monthId = monthDisplay ? this.monthNameToId[monthDisplay] : null;
    const year = Number(yearVal);
    const budgetAllocated = Number(budgetVal);
    const hoursPlanned = Number(hoursVal);

    let rowValid = true;

    // Cell-level validations with highlighting and specific reasons
    if (!projectDisplay || !projectCode) {
      invalidCells.push({ rowIndex, colIndex: 1, reason: 'Project name is required and must be a valid selection.' }); // A: Project
      rowValid = false;
    }
    if (!employeeDisplay || !employeeCode) {
      invalidCells.push({ rowIndex, colIndex: 2, reason: 'Employee name is required and must be a valid selection.' }); // B: Employee
      rowValid = false;
    }
    if (isNaN(year) || year < 1000 || year > 4050 || year !== Math.floor(year)) {
      invalidCells.push({ rowIndex, colIndex: 3, reason: 'Year must be a valid integer between 1000 and 4050.' }); // C: Year
      rowValid = false;
    }
    if (!monthDisplay || monthId === null || monthId === undefined) {
      invalidCells.push({ rowIndex, colIndex: 4, reason: 'Month is required and must be a valid selection.' }); // D: Month
      rowValid = false;
    }
    if (isNaN(budgetAllocated) || budgetAllocated <= 0) {
      invalidCells.push({ rowIndex, colIndex: 5, reason: 'Budget must be a number greater than 0.' }); // E: Budget
      rowValid = false;
    }
    if (isNaN(hoursPlanned) || hoursPlanned < 0 || hoursPlanned > 999) {
      invalidCells.push({ rowIndex, colIndex: 6, reason: 'Hours must be a number between 0 and 999.' }); // F: Hours
      rowValid = false;
    }
    if (!statusName || !this.statusNameToValue[statusName]) {
      invalidCells.push({ rowIndex, colIndex: 8, reason: 'Status is required and must be a valid selection.' }); // H: Status
      rowValid = false;
    }
    if (comments.length > 200) {
      invalidCells.push({ rowIndex, colIndex: 9, reason: 'Comments must not exceed 200 characters.' }); // I: Comments
      rowValid = false;
    }

    if (!rowValid) {
      invalidRows.push(rowIndex);
      return;
    }

    upserts.push({
      projectCode:projectCode!,
      employeeCode:employeeCode!,
      year,
      monthId: Number(monthId),
      budgetAllocated,
      hoursPlanned,
      statusName,
      comments
    });
  });

  console.log('Prepared upserts:', upserts); // For debugging
  console.log('Invalid rows (skipped):', invalidRows); // For debugging

if (invalidCells.length > 0) {
  const widget = (this.spreadsheet as any).spreadsheetWidget;
  const sheet = widget.activeSheet();

  // Highlight only invalid cells
  invalidCells.forEach(({ rowIndex, colIndex }) => {
    const colLetter = this.getColumnLetter(colIndex);
    const cell = sheet.range(`${colLetter}${rowIndex}`);
    if (cell) {
      cell.background('#f15663ff');
      cell.color('#721c24');
    }
  });

    // Specific alert: Group unique issues by row
    const issuesByRow: Record<number, Set<string>> = {}; // Use Set to avoid duplicates
    invalidCells.forEach(({ rowIndex, reason }) => {
      if (!issuesByRow[rowIndex]) issuesByRow[rowIndex] = new Set();
      issuesByRow[rowIndex].add(reason);
    });

    let alertMsg = 'Please correct the highlighted fields before submitting.\n\n';
    alertMsg += `Issues found in rows: ${[...new Set(invalidRows)].sort((a, b) => a - b).join(', ')}.\n\n`;
    alertMsg += 'Specific issues:\n';

    Object.entries(issuesByRow).sort(([a], [b]) => Number(a) - Number(b)).forEach(([row, reasons]) => {
      alertMsg += `\nRow ${row}:\n`;
      reasons.forEach(reason => {
        alertMsg += `• ${reason}\n`;
      });
    });

    alert(alertMsg);
    return;
  }

  if (upserts.length) {
    this.service.bulkUpsert(upserts).subscribe((affected) => {
      let msg = `Saved ${affected} rows successfully!`;
      if (invalidRows.length > 0) {
        msg += ` Skipped ${invalidRows.length} invalid rows (e.g., invalid dropdown values or numbers).`;
      }
      if (affected === 0) {
        msg += ' No changes were made (e.g., no updates needed).';
      }
      alert(msg);
      // Reset formatting state and reapply
      this.lastFormattedRowCount = 0;
      setTimeout(() => {
        this.attachFormulasAndFormatting();
      }, 100);
      this.loadData(); // Reload to reflect changes
    }, err => {
      console.error('Save error', err);
      alert('Save failed. Check console for details.');
    });
  } else {
    let msg = 'No valid rows to save.';
    if (invalidRows.length > 0) {
      msg += ` All ${invalidRows.length} data rows had issues (e.g., invalid dropdown selections from paste, non-numeric values).`;
    }
    alert(msg);
  }
}


// Add this helper method to your component class (SpreadsheetWrapperComponent)
private getColumnLetter(col: number): string {
  let letter = '';
  let tempCol = col;
  while (tempCol > 0) {
    tempCol--;
    letter = String.fromCharCode(65 + (tempCol % 26)) + letter;
    tempCol = Math.floor(tempCol / 26);
  }
  return letter;
}


async deleteSelected() {
    const selectedIndices = this.selectedRows();  // Use stored selection
    console.log('Selected row indices from event:', selectedIndices);

    const idsToDelete = selectedIndices
      .map(r => this.rowToId[r])
      .filter(id => !!id);

    console.log('Resolved IDs to delete:', idsToDelete);
    
    if (!idsToDelete.length) {
      alert('No valid rows selected to delete.');
      return;
    }

    if (!confirm(`Delete ${idsToDelete.length} plan(s)?`)) return;

    try {
      await firstValueFrom(this.service.bulkDelete(idsToDelete));
      this.loadData();  // Reload to reflect changes
      this.selectedRows.set([]);  // Clear selection after delete
    } catch (err) {
      console.error('Error deleting plans:', err);
      alert('Failed to delete selected plans.');
    }
  }

exportToPDF() {
  // Fetch the spreadsheet workbook data
  const workbook = (this.spreadsheet as any).spreadsheetWidget.toJSON();
  const rows = workbook.sheets?.[0]?.rows ?? [];

  // Remove header row
  const dataRows = rows.slice(1);

  if (!dataRows.length) {
    alert('No data available to export');
    return;
  }

  // Define headers for PDF
  const headers = [
    ['Project', 'Employee', 'Year', 'Month', 'Budget', 'Hours', 'Cost', 'Status', 'Comments']
  ];

  // Map spreadsheet rows into plain arrays
  const body = dataRows.map((r: any) => {
    const c = r.cells ?? [];
    return [
      c[0]?.value ?? '',
      c[1]?.value ?? '',
      c[2]?.value ?? '',
      c[3]?.value ?? '',
      c[4]?.value ?? '',
      c[5]?.value ?? '',
      c[6]?.value ?? '',
      c[7]?.value ?? '',
      c[8]?.value ?? ''
    ];
  });

  // Create the PDF
  const doc = new jsPDF('l', 'mm', 'a4'); // landscape for wider tables

  autoTable(doc, {
    head: headers,
    body,
    startY: 20,
    styles: {
      fontSize: 8,
      cellPadding: 3,
    },
    headStyles: {
      fillColor: [41, 128, 185],
      textColor: 255,
      halign: 'center'
    },
    columnStyles: {
      0: { cellWidth: 30 }, // Project
      1: { cellWidth: 30 }, // Employee
      2: { cellWidth: 20 }, // Year
      3: { cellWidth: 20 }, // Month
      4: { cellWidth: 25 }, // Budget
      5: { cellWidth: 25 }, // Hours
      6: { cellWidth: 25 }, // Cost
      7: { cellWidth: 30 }, // Status
      8: { cellWidth: 70 }, // Comments (wider)
    }
  });

  doc.save('BudgetPlans.pdf');
}


importExcel(event: any) {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target!.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const firstSheet = workbook.Sheets[firstSheetName];
        const json = XLSX.utils.sheet_to_json(firstSheet, { header: 1 }) as any[][];

        if (json.length < 1) {
          alert('Excel file is empty.');
          return;
        }

        // Assume first row is headers; skip if they don't match expected
        const headers = json[0] as string[];
        const dataRows = json.slice(1);

        // Expected column order: Project, Employee, Year, Month, Budget, Hours, Status, Comments
        // (Cost can be ignored or calculated)
        const colMap: Record<string, number> = {};
        headers.forEach((h, idx) => {
          const cleanH = h?.toString().trim().toLowerCase();
          if (cleanH === 'project' || cleanH === 'project name' || cleanH === 'projectcode') colMap['project'] = idx;
          else if (cleanH === 'employee' || cleanH === 'employee name' || cleanH === 'employeecode') colMap['employee'] = idx;
          else if (cleanH === 'year') colMap['year'] = idx;
          else if (cleanH === 'month' || cleanH === 'month name') colMap['month'] = idx;
          else if (cleanH === 'budget' || cleanH === 'budgetallocated') colMap['budget'] = idx;
          else if (cleanH === 'hours' || cleanH === 'hoursplanned') colMap['hours'] = idx;
          else if (cleanH === 'status' || cleanH === 'statusname') colMap['status'] = idx;
          else if (cleanH === 'comments') colMap['comments'] = idx;
        });

        const upserts: BudgetPlanUpsert[] = [];
        const invalidRows: number[] = [];

        dataRows.forEach((row, rowIdx) => {
          const projectDisplay = row[colMap['project'] ?? 0]?.toString().trim();
          const employeeDisplay = row[colMap['employee'] ?? 1]?.toString().trim();
          const yearVal = row[colMap['year'] ?? 2];
          const monthDisplay = row[colMap['month'] ?? 3]?.toString().trim();
          const budgetVal = row[colMap['budget'] ?? 4];
          const hoursVal = row[colMap['hours'] ?? 5];
          const statusName = row[colMap['status'] ?? 7]?.toString().trim();
          const comments = row[colMap['comments'] ?? 8]?.toString().trim() ?? '';

          // Similar validation/mapping as saveData
          const projectCode = projectDisplay ? this.projectNameToCode[projectDisplay] : null;
          const employeeCode = employeeDisplay ? this.employeeNameToCode[employeeDisplay] : null;
          const monthId = monthDisplay ? this.monthNameToId[monthDisplay] : null;
          const year = Number(yearVal);
          const budgetAllocated = Number(budgetVal);
          const hoursPlanned = Number(hoursVal);

          if (!projectCode || !employeeCode || monthId === null || monthId === undefined ||
              isNaN(year) || year < 1000 || year > 4050 || year !== Math.floor(year) ||
              isNaN(budgetAllocated) || budgetAllocated <= 0 ||
              isNaN(hoursPlanned) || hoursPlanned < 0 || hoursPlanned > 999 ||
              !statusName || !this.statusNameToValue[statusName] ||
              comments.length > 500) {
            invalidRows.push(rowIdx + 2);
            return;
          }

          upserts.push({
            projectCode,
            employeeCode,
            year,
            monthId: Number(monthId),
            budgetAllocated,
            hoursPlanned,
            statusName,
            comments
          });
        });

        console.log('Prepared import upserts:', upserts);
        console.log('Invalid import rows (skipped):', invalidRows);

        if (upserts.length) {
          this.service.bulkUpsert(upserts).subscribe((affected) => {
            let msg = `Imported and saved ${affected} rows successfully!`;
            if (invalidRows.length > 0) {
              msg += ` Skipped ${invalidRows.length} invalid rows.`;
            }
            if (affected === 0) {
              msg += ' No changes were made.';
            }
            alert(msg);
            this.loadData(); // Reload to reflect changes
          }, err => {
            console.error('Import error', err);
            alert('Import failed. Check console for details.');
          });
        } else {
          let msg = 'No valid rows to import.';
          if (invalidRows.length > 0) {
            msg += ` All ${invalidRows.length} rows had issues (e.g., invalid values or column mismatch). Ensure columns match: Project, Employee, Year, Month, Budget, Hours, (Cost), Status, Comments.`;
          }
          alert(msg);
        }
      } catch (err) {
        console.error('Excel parse error:', err);
        alert('Failed to parse Excel file. Ensure it\'s a valid .xlsx/.xls file.');
      }
    };
    reader.readAsArrayBuffer(file);

    // Reset file input
    event.target.value = '';
  }


}