// Enums
export enum DataTypes {
    String = 0,
    Number = 1,
    Date = 2,
    Boolean = 3,
    Decimal = 4
}

export enum ConfigType {
    Salary = 0,
    Insurance = 1
}

export enum DepartmentId {
    All = 0,
    DepartmentA = 1,
    DepartmentB = 2
}

// Types
export interface ExcelConfigDetail {
    id?: number;
    configId?: number;
    fieldName: string;
    displayName: string;
    columnPosition: number;
    rowPosition: number;
    sheetName: string;
    dataType: DataTypes;
    isRequired: boolean;
}

export interface ExcelConfig {
    id: number;
    templateFileName: string;
    configName: string;
    departmentId: number;
    configType: ConfigType;
    details?: ExcelConfigDetail[];
    acctions: string;
}

export const API_BASE_URL = 'https://localhost:7034';

export interface HeaderMapping {
    col: number;
    originalValue: string;
    displayName: string;
    rowIndex: number;
    sheet: string;
}

export interface Cell {
    row: number;
    col: number;
    sheet: string;
}

export interface CellError extends Cell {
    err?: string;
    index?: number;
}

export interface DataStartCell extends Cell {
    field?: string;
}

export interface Table {
    tableName: string;
    fields: Field[];
}

export interface Field {
    fieldName: string;
    nameDisplay: string;
    type: DataTypes;
    isSelected?: boolean;
    isRequired?: boolean;
}

export interface TryCastResult<T> {
    success: boolean;
    value: T | null;
    error?: string;
}

export interface SubmitResult<T = any> {
    isSuccess: boolean;
    data: T;
    cellsErr: CellError[];
}

export type Step = 'select_mode' | 'select_headers' | 'set_row_start' | 'select_data_start' | 'configure';

export const Tables: Field[] = [
    { fieldName: 'fullName', nameDisplay: 'Họ và tên', type: 0, isRequired: false },
    { fieldName: 'ctvCode', nameDisplay: 'Mã CTV', type: 0, isRequired: false },

    { fieldName: 'firstContractStartDateT9_2024', nameDisplay: 'Ngày bắt đầu hợp đồng lần (T9/2024)', type: 2, isRequired: false },
    { fieldName: 'contractStartDate', nameDisplay: 'Ngày bắt đầu hợp đồng', type: 2, isRequired: false },

    { fieldName: 'organization', nameDisplay: 'Đơn vị', type: 0, isRequired: false },
    { fieldName: 'jobPosition', nameDisplay: 'Vị trí công việc', type: 0, isRequired: false },

    { fieldName: 'actualWorkingDays', nameDisplay: 'Ngày công thực tế', type: 1, isRequired: false },
    { fieldName: 'leaveDays', nameDisplay: 'Ngày công phép', type: 1, isRequired: false },
    { fieldName: 'holidayDays', nameDisplay: 'Ngày công lễ', type: 1, isRequired: false },
    { fieldName: 'nightShiftDays', nameDisplay: 'Ngày công ca đêm', type: 1, isRequired: false },
    { fieldName: 'policyLeaveDays', nameDisplay: 'Nghỉ chế độ', type: 1, isRequired: false },
    { fieldName: 'bhxhLeaveDays', nameDisplay: 'Nghỉ BHXH', type: 1, isRequired: false },
    { fieldName: 'unpaidLeaveDays', nameDisplay: 'Ngày nghỉ không lương', type: 1, isRequired: false },

    { fieldName: 'vtcvSalaryWorkingDays', nameDisplay: 'Tổng công tính lương VTCV', type: 1, isRequired: false },
    { fieldName: 'performanceSalaryWorkingDays', nameDisplay: 'Tổng công tính lương hiệu quả', type: 1, isRequired: false },
    { fieldName: 'actualSalaryWorkingDaysHidden', nameDisplay: 'Tổng công tính lương thực tế ẩn', type: 1, isRequired: false },

    { fieldName: 'nightShiftWorkingDays', nameDisplay: 'Ngày công ca đêm', type: 1, isRequired: false },
    { fieldName: 'holidayDutyWorkingDays', nameDisplay: 'Ngày công trực ca lễ tết', type: 1, isRequired: false },
    { fieldName: 'standardWorkingDaysOfMonth', nameDisplay: 'Ngày công chuẩn của tháng', type: 1, isRequired: false },

    { fieldName: 'bhxhBaseSalary', nameDisplay: 'Mức lương làm căn cứ đóng BHXH', type: 1, isRequired: false },
    { fieldName: 'vtcvSalary', nameDisplay: 'Tiền lương VTCV', type: 1, isRequired: false },
    { fieldName: 'workCompletionRate', nameDisplay: 'Tỉ lệ hoàn thành công việc', type: 1, isRequired: false },
    { fieldName: 'performanceSalary', nameDisplay: 'Lương hiệu quả', type: 1, isRequired: false },
    { fieldName: 'nightAndHolidaySalary', nameDisplay: 'Lương ca đêm và trực ca lễ tết', type: 1, isRequired: false },

    { fieldName: 'totalVtcvAndPerformanceSalary', nameDisplay: 'Tổng lương VTCV và hiệu quả', type: 1, isRequired: false },
    { fieldName: 'agreedSalaryColumn', nameDisplay: 'Cột lương thỏa thuận trả cho người lao động', type: 1, isRequired: false },
    { fieldName: 'salaryArrears', nameDisplay: 'Truy lĩnh tiền lương', type: 1, isRequired: false },
];
