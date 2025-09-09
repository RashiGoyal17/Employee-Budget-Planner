import { HttpClient, HttpParams } from '@angular/common/http';
import { Injectable } from '@angular/core';
import { Observable } from 'rxjs';
import { BudgetPlan, BudgetPlanUpsert, PagedResult } from '../models/BudgetPlanModel';

@Injectable({
  providedIn: 'root'
})
export class BudgetPlanService {

  private readonly apiRoot = 'https://localhost:7225/api';

  constructor(private http: HttpClient) {}

  getPlans(query: {
    projectCode?: string;
    employeeCode?: string;
    year?: number;
    monthId?: number;
    page?: number;
    pageSize?: number;
  }): Observable<PagedResult<BudgetPlan>> {
    let params = new HttpParams();
    Object.entries(query).forEach(([k, v]) => {
      if (v !== undefined && v !== null) params = params.set(k, v.toString());
    });
    return this.http.get<PagedResult<BudgetPlan>>(`${this.apiRoot}/budgetplans`, { params });
  }

  bulkUpsert(items: BudgetPlanUpsert[]): Observable<void> {
    return this.http.post<void>(`${this.apiRoot}/budgetplans/bulk`, items);
  }

  
  getProjects()   { return this.http.get<{ code: string; name: string }[]>(`${this.apiRoot}/projects`); }
  getEmployees()  { return this.http.get<{ employeeCode: string; name: string }[]>(`${this.apiRoot}/employees`); }
  getStatuses()   { return this.http.get<{ name: string }[]>(`${this.apiRoot}/statuses`); }
  getMonths()     { return this.http.get<{ monthId: number; name: string }[]>(`${this.apiRoot}/months`); }
  bulkDelete(ids: number[]): Observable<void> {
  return this.http.post<void>(`${this.apiRoot}/budgetplans/bulk-delete`, ids);
}

}