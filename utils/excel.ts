import { CellError, DataTypes, ExcelConfigDetail, Field, SubmitResult, TryCastResult } from '../types/excel-type';
import * as XLSX from 'xlsx';

export const isEmptyValue = (value: any): boolean => {
    return value === null ||
        value === undefined ||
        value === "" ||
        (typeof value === 'string' && value.trim() === "");
};

export const tryCast = (
    value: any,
    type: DataTypes
): TryCastResult<any> => {

    if (value === undefined || value === null || value === '') {
        return { success: true, value: null };
    }

    try {
        switch (type) {
            case DataTypes.Number: {
                const num = Number(value);
                if (isNaN(num)) {
                    return {
                        success: false,
                        value: null,
                        error: `"${value}" không phải là số`
                    };
                }
                return { success: true, value: num };
            }

            case DataTypes.Boolean: {
                if (typeof value === 'boolean') {
                    return { success: true, value };
                }

                if (value === 1 || value === '1' || value === 'true') {
                    return { success: true, value: true };
                }

                if (value === 0 || value === '0' || value === 'false') {
                    return { success: true, value: false };
                }

                return {
                    success: false,
                    value: null,
                    error: `"${value}" không phải boolean`
                };
            }

            case DataTypes.Date: {
                let date: Date | null = null;

                // Xử lý Excel serial date (nếu value là số nguyên dương)
                if (typeof value === 'number' && Number.isInteger(value) && value > 0) {
                    // Excel serial date bắt đầu từ 1900-01-01 (giá trị 1)
                    // Công thức: Date(1899, 11, 30) + value (vì Excel có bug ở 1900 không nhuận, nhưng ta dùng offset chuẩn)
                    const excelBaseDate = new Date(1899, 11, 30); // Base cho serial
                    date = new Date(excelBaseDate.getTime() + value * 86400000); // 86400000 ms = 1 ngày
                }
                // Xử lý string dạng dd/MM/yyyy (hoặc dd-MM-yyyy)
                else if (typeof value === 'string') {
                    const ddmmyyyyRegex = /^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/;
                    const match = value.match(ddmmyyyyRegex);
                    if (match) {
                        const day = parseInt(match[1], 10);
                        const month = parseInt(match[2], 10);
                        const year = parseInt(match[3], 10);
                        date = new Date(year, month - 1, day);
                    }
                }

                // Fallback: Sử dụng new Date(value) cho các định dạng khác (ISO, MM/dd/yyyy, etc.)
                if (!date || isNaN(date.getTime())) {
                    date = new Date(value);
                }

                if (isNaN(date.getTime())) {
                    return {
                        success: false,
                        value: null,
                        error: `"${value}" không phải ngày hợp lệ`
                    };
                }

                return {
                    success: true,
                    value: date.toISOString()
                };
            }

            case DataTypes.String:
            default:
                return {
                    success: true,
                    value: value.toString()
                };
        }
    } catch (err) {
        return {
            success: false,
            value: null,
            error: (err as Error).message
        };
    }
};

export const mappingDataType = (type: DataTypes) : string =>
{
    switch (type){
        case DataTypes.Boolean:
            return 'bool';
        case DataTypes.Date:
            return 'date'
        case DataTypes.Decimal:
        case DataTypes.Number:
            return 'number';
        default:
            return 'string';
    }
}

export const extractDataWithConfig = (
        workbook: XLSX.WorkBook,
        dataStartCells: ExcelConfigDetail[],
        fields: Field[]
    ): SubmitResult<Record<string, any>[]> => {
        const result: Record<string, any>[] = [];
        let isSuccess: boolean = true;
        let cellsErr: CellError[] = [];
        const columnData: any[][] = dataStartCells.map(cfg => {
            const worksheet = workbook.Sheets[cfg.sheetName];
            if (!worksheet) return [];

            const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: null }) as any[][];
            const data: any[] = [];
            for (let i = cfg.rowPosition; i < sheetData.length; i++) {
                const row = sheetData[i] || [];
                data.push(row[cfg.columnPosition] ?? null);
            }
            return data;
        });

        const maxRows = Math.max(...columnData.map(col => col.length), 0);

        for (let rowIdx = 0; rowIdx < maxRows; rowIdx++) {
            const dataRow: Record<string, any> = {};
            let allEmpty = true;

            for (let colIdx = 0; colIdx < dataStartCells.length; colIdx++) {
                const value = columnData[colIdx][rowIdx] ?? null;
                const fieldName = dataStartCells[colIdx].fieldName;
                const field = fields.find(f => f.fieldName === fieldName);

                let res = tryCast(value, field?.type ?? DataTypes.String);
                if (!res.success) {
                    const startCol = dataStartCells[colIdx].columnPosition;
                    const startRow = dataStartCells[colIdx].rowPosition;
                    const sheet = dataStartCells[colIdx].sheetName;
                    const mess = res.error ?? '';
                    cellsErr.push({
                        col: startCol,
                        row: startRow + rowIdx,
                        sheet: sheet,
                        index: colIdx,
                        err: mess
                    });
                    isSuccess = false;
                }
                dataRow[fieldName] = res.success ? res.value : res.error;
                if (!isEmptyValue(value)) {
                    allEmpty = false;
                }
            }

            if (allEmpty) break;
            result.push(dataRow);
        }

        return { isSuccess, data: result, cellsErr };
    };