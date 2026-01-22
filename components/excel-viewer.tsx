import { Tabs, Tab } from "@heroui/react";
import { CircleCheck } from "lucide-react";
import React, { useState, useRef, useMemo, useCallback } from 'react';
import * as XLSX from 'xlsx';

interface ExcelViewerProps {
    workbook: XLSX.WorkBook;
    onCellClick?: (rowIdx: number, colIdx: number, sheetName: string) => void;
    getCellClassName?: (rowIdx: number, colIdx: number, sheetName: string) => string;
    selectedSheet?: string;
    onSheetChange?: (sheetName: string) => void;
    highlightedCells?: Set<string>;
    readOnly?: boolean;
    sheetConfigured: Set<string>;
}

export function ExcelViewer({
    workbook,
    onCellClick,
    getCellClassName,
    selectedSheet: externalSelectedSheet,
    onSheetChange,
    highlightedCells = new Set(),
    readOnly = false,
    sheetConfigured = new Set()
}: ExcelViewerProps) {
    const [internalSelectedSheet, setInternalSelectedSheet] = useState(workbook.SheetNames[0]);
    const [scrollTop, setScrollTop] = useState(0);
    const tableRef = useRef<HTMLDivElement>(null);

    const selectedSheet = externalSelectedSheet || internalSelectedSheet;
    const sheets = workbook.SheetNames;

    const rowHeight = 24;
    const colWidth = 120;
    const visibleRows = 20;

    const data = useMemo(() => {
        const worksheet = workbook.Sheets[selectedSheet];
        return XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' }) as (string | number | boolean | null)[][];
    }, [workbook, selectedSheet]);

    const handleSheetChange = (sheetName: string) => {
        if (onSheetChange) {
            onSheetChange(sheetName);
        } else {
            setInternalSelectedSheet(sheetName);
        }
        setScrollTop(0);
    };

    const isSheetConfigured = (sheetName: string): boolean => {
        return sheetConfigured?.has(sheetName) ?? false;
    }

    const handleCellClick = (rowIdx: number, colIdx: number) => {
        if (!readOnly && onCellClick) {
            onCellClick(rowIdx, colIdx, selectedSheet);
        }
    };

    const maxCols = Math.max(...data.map(row => row.length), 0);
    const startRow = Math.floor(scrollTop / rowHeight);
    const endRow = Math.min(startRow + visibleRows, data.length);

    const handleScroll = useCallback((e: React.UIEvent<HTMLDivElement>) => {
        setScrollTop(e.currentTarget.scrollTop);
    }, []);

    const defaultCellClassName = (rowIdx: number, colIdx: number) => {
        const cellKey = `${rowIdx}-${colIdx}-${selectedSheet}`;
        if (highlightedCells.has(cellKey)) {
            return 'bg-green-200 font-bold border-2 border-green-500';
        }
        return readOnly ? 'bg-white' : 'bg-white hover:bg-gray-100 cursor-pointer';
    };

    const excelColName = (col: number): string => {
        let name = '';
        while (col >= 0) {
            name = String.fromCharCode((col % 26) + 65) + name;
            col = Math.floor(col / 26) - 1;
        }
        return name;
    };

    const [hoveredCell, setHoveredCell] = useState<{ rowIdx: number; colIdx: number; text: string; position: { top: number; left: number } } | null>(null);

    const handleMouseEnter = (event: React.MouseEvent<HTMLTableCellElement>, rowIdx: number, colIdx: number, text: any) => {
        const cellText = text ? text.toString() : '';
        if (cellText.length > 13) {
            const rect = event.currentTarget.getBoundingClientRect();
            let left = rect.right + 10;
            if (left + 300 > window.innerWidth) {
                left = rect.left - 310;
            }
            setHoveredCell({
                rowIdx,
                colIdx,
                text: cellText,
                position: { top: rect.top, left }
            });
        }
    };

    const handleMouseLeave = () => {
        setHoveredCell(null);
    };

    return (
        <div className="border-2 border-blue-200 rounded-lg p-0.5 shadow-md overflow-hidden">
            <div className="flex border-b overflow-x-auto">
                <Tabs
                    key='tab-underlined'
                    variant='solid'
                    selectedKey={selectedSheet}
                    onSelectionChange={(e) => handleSheetChange(e.toString())}
                >
                    {sheets.map((sheet) => (
                        <Tab
                            key={sheet}
                            title={
                                <div className="flex items-center gap-1">
                                    <span>{sheet}</span>
                                    {isSheetConfigured(sheet) && <CircleCheck className="text-emerald-500" />}
                                </div>
                            }
                        />
                    ))}
                </Tabs>
            </div>

            <div className="overflow-hidden bg-gray-100 border-b border-gray-300">
                <div style={{ width: maxCols * colWidth + 50, minWidth: '100%' }}>
                    <table className="border-collapse">
                        <thead>
                            <tr style={{ height: rowHeight }}>
                                <th className="border border-gray-300 px-2 py-1 text-xs font-semibold text-gray-700 bg-gray-100" style={{ width: '50px' }}>
                                    #
                                </th>
                                {Array.from({ length: maxCols }).map((_, colIdx) => (
                                    <th
                                        key={colIdx}
                                        className="border border-gray-300 px-2 py-1 text-xs font-semibold text-gray-700 bg-gray-100"
                                        style={{ minWidth: colWidth, maxWidth: colWidth }}
                                    >
                                        {excelColName(colIdx)}
                                    </th>
                                ))}
                            </tr>
                        </thead>
                    </table>
                </div>
            </div>

            <div
                ref={tableRef}
                className="overflow-auto relative bg-white"
                onScroll={handleScroll}
                style={{ height: '564px', width: '100%' }}
            >
                <div style={{
                    height: data.length * rowHeight,
                    width: maxCols * colWidth + 50,
                    position: 'relative'
                }}>
                    <table className="w-full absolute border-collapse" style={{
                        top: startRow * rowHeight,
                        left: 0
                    }}>
                        <tbody>
                            {data.slice(startRow, endRow).map((row, relRowIdx) => {
                                const rowIdx = startRow + relRowIdx;
                                return (
                                    <tr key={rowIdx} style={{ height: rowHeight }}>
                                        <td className="border border-gray-300 text-gray-600 px-2 py-1 text-xs bg-gray-100 font-semibold text-center" style={{ width: '50px' }}>
                                            {rowIdx + 1}
                                        </td>
                                        {Array.from({ length: maxCols }).map((_, colIdx) => {
                                            const cellValue = row[colIdx] || '';
                                            const cellClassName = getCellClassName
                                                ? getCellClassName(rowIdx, colIdx, selectedSheet)
                                                : defaultCellClassName(rowIdx, colIdx);

                                            return (
                                                <td
                                                    key={colIdx}
                                                    onClick={() => handleCellClick(rowIdx, colIdx)}
                                                    onMouseEnter={(e) => handleMouseEnter(e, rowIdx, colIdx, cellValue)}
                                                    onMouseLeave={handleMouseLeave}
                                                    className={`border text-gray-600 border-gray-300 px-3 py-2 text-sm ${cellClassName}`}
                                                    style={{
                                                        minWidth: colWidth,
                                                        maxWidth: colWidth,
                                                        overflow: 'hidden',
                                                        textOverflow: 'ellipsis',
                                                        whiteSpace: 'nowrap'
                                                    }}
                                                >
                                                    {cellValue}
                                                </td>
                                            );
                                        })}
                                    </tr>
                                );
                            })}
                        </tbody>
                    </table>
                </div>

                {hoveredCell && (
                    <div
                        className="fixed z-50 bg-gray-100 text-gray-500 border rounded border-gray-300 shadow-lg p-2 max-w-xs whitespace-normal overflow-auto max-h-40 pointer-events-none"
                        style={{
                            top: `${hoveredCell.position.top}px`,
                            left: `${hoveredCell.position.left}px`,
                        }}
                    >
                        {hoveredCell.text}
                    </div>
                )}
            </div>
        </div>
    );
}