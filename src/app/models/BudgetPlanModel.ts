export interface BudgetPlan{
    budgetPlanId: number;
    projectCode: string;
    employeeCode:string;
    year:number;
    monthId:number;
    budgetAllocated: number;
    hoursPlanned: number;
    cost: number;
    status : string;
    comments?: string;
}

export interface BudgetPlanUpsert{
    projectCode: string;
    employeeCode: string;
    year: number;
    monthId: number;
    budgetAllocated: number;
    hoursPlanned: number;
    statusName: string;
    comments?:string;
}

export interface PagedResult<T>{
    total: number;
    items: T[];
}