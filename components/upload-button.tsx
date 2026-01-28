'use client'

import React, { useEffect, useState } from 'react';
import * as XLSX from 'xlsx';
import { Upload, FileSpreadsheet, CheckCircle, X, Check, Download } from 'lucide-react';
import { ExcelViewer } from '@/components/excel-viewer';
import { Input } from '@heroui/input';
import { Button } from '@heroui/button';
import { Table, TableBody, TableCell, TableColumn, TableHeader, TableRow, useDisclosure, Modal, ModalBody, ModalContent, ModalFooter, ModalHeader } from '@heroui/react';
import { API_BASE_URL, CellError, ConfigType, DataStartCell, DataTypes, ExcelConfig, ExcelConfigDetail, Field, SubmitResult, Tables } from '@/types/excel-type';
import { extractDataWithConfig, isEmptyValue, tryCast } from '@/utils/excel';

interface ExcelImporterProps {
    departmentId: number;
    configType: number;
    name?: string;
    color?: "default" | "primary" | "secondary" | "success" | "warning" | "danger" | undefined;
    size?: "sm" | "md" | "lg" | undefined;
}

export default function UploadButton({ departmentId, configType, name, color, size }: ExcelImporterProps) {
    const [file, setFile] = useState<File | null>(null);
    const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
    const [selectedSheet, setSelectedSheet] = useState('');
    const [fields, setFields] = useState<Field[]>([]);
    const [dataStartCells, setDataStartCells] = useState<ExcelConfigDetail[]>([]);
    const [extractedData, setExtractedData] = useState<Record<string, any>[]>([]);
    const [previewData, setPreviewData] = useState<Record<string, any>[]>([]);
    const { isOpen: isImportOpen, onOpen: onImportOpen, onOpenChange: onImportOpenChange } = useDisclosure();
    const { isOpen: isPreviewOpen, onOpen: onPreviewOpen, onOpenChange: onPreviewOpenChange } = useDisclosure();
    const [cellError, setCellError] = useState<CellError[]>([]);
    const [loading, setLoading] = useState(false);
    const [sheetsConfigured, setSheetsConfigured] = useState<Set<string>>(new Set());

    const fetchConfigDetails = async (configId: number) => {
        setLoading(true);
        try {
            const res = await fetch(`https://localhost:7034/excel-config/${configId}/details`);
            if (!res.ok) throw new Error('Lỗi fetch config details');
            const details: ExcelConfigDetail[] = await res.json();
            setFields(Tables);
            setDataStartCells(details);

            setSheetsConfigured(new Set(details.map(m => m.sheetName)));
        } catch (error) {
            alert('Lỗi fetch config: ' + (error as Error).message);
        } finally {
            setLoading(false);
        }
    };

    const fetchConfig = async (): Promise<ExcelConfig> => {
        setLoading(true);
        try {
            const res = await fetch(`${API_BASE_URL}/excel-config/config?departmentId=${departmentId}&configType=${configType}`);
            if (!res.ok) throw new Error('Lỗi fetch config details');
            const config: ExcelConfig = await res.json();
            return config;
        } catch (error) {
            alert('Lỗi fetch config: ' + (error as Error).message);
            throw error;
        } finally {
            setLoading(false);
        }
    };

    const filterVisibleWorkbook = (wb: XLSX.WorkBook): XLSX.WorkBook => {
        const visibleSheetNames = wb.SheetNames.filter(name => {
            const sheetMeta = wb.Workbook?.Sheets?.find(s => s.name === name);
            return !sheetMeta || sheetMeta.Hidden === 0;
        });

        const newWb = XLSX.utils.book_new();

        visibleSheetNames.forEach(name => {
            XLSX.utils.book_append_sheet(newWb, wb.Sheets[name], name);
        });

        return newWb;
    };

    const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
        const uploadedFile = e.target.files?.[0];
        if (!uploadedFile) return;

        setFile(uploadedFile);
        const reader = new FileReader();

        reader.onload = (event) => {
            try {
                const wbRaw = XLSX.read(event.target?.result, { type: 'binary' });
                const wb = filterVisibleWorkbook(wbRaw);
                setWorkbook(wb);
                const firstSheet = wb.SheetNames[0];
                setSelectedSheet(firstSheet);
            } catch (error) {
                alert('Lỗi khi đọc file Excel: ' + (error as Error).message);
            }
        };

        reader.readAsBinaryString(uploadedFile);
    };

    const checkRequiredFields = () => {
        const mappedFields = new Set(dataStartCells.map(d => d.fieldName));
        return fields.filter(f => f.isRequired).every(f => mappedFields.has(f.fieldName));
    };

    const handleExtract = () => {
        if (!workbook || fields.length === 0) return;

        if (!checkRequiredFields()) {
            alert('Config thiếu trường bắt buộc!');
            return;
        }

        const result = extractDataWithConfig(workbook, dataStartCells, fields);
        setExtractedData(result.data);
        setPreviewData(result.data.map((d, i) => ({
            key: `key_${i}`,
            stt: i + 1,
            ...d
        })));

        if (!result.isSuccess) {
            setCellError(result.cellsErr);
        } else {
            setCellError([]);
        }

        onPreviewOpen();
    };

    const getCellStyle = (rowIdx: number, colIdx: number, sheet: string) => {
        if (cellError.some(cell => cell.row === rowIdx && cell.col === colIdx && cell.sheet === sheet)) {
            return 'bg-red-500 text-white font-semibold border-2 border-blue-500';
        }
        if (dataStartCells.some(cell => cell.rowPosition === rowIdx && cell.columnPosition === colIdx && cell.sheetName === sheet)) {
            return 'bg-blue-200 border-2 border-gray-400';
        }
        return 'bg-white';
    };

    const resetFile = () => {
        setFile(null);
        setWorkbook(null);
        setSelectedSheet('');
        setCellError([]);
        setPreviewData([]);
        setFields([]);
        setDataStartCells([]);
    };

    const loadConfig = async () => {
        const config = await fetchConfig();
        if (config?.id) {
            await fetchConfigDetails(config.id);
        }
    };

    useEffect(() => {
        if (!workbook) return;

        loadConfig();
    }, [workbook]);

    return (
        <>
            <Button
                onPress={onImportOpen}
                color={color}
                size={size}
            >
                <Upload className="mx-auto" size={24} /> {name}
            </Button>

            <Modal
                size='5xl'
                isOpen={isImportOpen}
                onOpenChange={onImportOpenChange}
                onClose={() => resetFile()}
                isDismissable={false}
                isKeyboardDismissDisabled={true}
            >
                <ModalContent>
                    {(onClose) => (
                        <>
                            <ModalHeader>Import Excel</ModalHeader>
                            <ModalBody>
                                {!file ? (
                                    <div className="border-2 border-dashed border-gray-300 rounded-lg p-12 text-center hover:border-blue-400 transition-colors">
                                        <Upload className="mx-auto mb-4" size={48} />
                                        <label className="cursor-pointer">
                                            <span className="text-lg hover:text-blue-600">
                                                Nhấp để chọn file Excel
                                            </span>
                                            <Input
                                                type="file"
                                                accept=".xlsx,.xls"
                                                onChange={handleFileChange}
                                                className="hidden"
                                            />
                                        </label>
                                        <p className="text-sm text-gray-500 mt-2">Hỗ trợ định dạng .xlsx và .xls</p>
                                    </div>
                                ) : (
                                    <div className="mb-6 flex items-center justify-between border-2 border-blue-200 rounded-lg shadow-md p-4">
                                        <div className="flex items-center gap-3">
                                            <CheckCircle className="text-green-600" />
                                            <div>
                                                <p className="font-semibold">{file.name}</p>
                                                <p className="text-sm">
                                                    {(file.size / 1024).toFixed(2)} KB
                                                </p>
                                            </div>
                                        </div>
                                        <Button
                                            onClick={resetFile}
                                            color='danger'
                                            startContent={<X size={16} />}
                                        >
                                            Xóa file
                                        </Button>
                                    </div>
                                )}

                                {workbook && (
                                    <ExcelViewer
                                        workbook={workbook}
                                        selectedSheet={selectedSheet}
                                        onSheetChange={setSelectedSheet}
                                        getCellClassName={getCellStyle}
                                        readOnly={true}
                                        sheetsConfigured={sheetsConfigured}
                                    />
                                )}

                                {loading && <p>Đang tải config...</p>}
                            </ModalBody>
                            <ModalFooter>
                                <Button color="danger" onPress={onClose}>
                                    Đóng
                                </Button>
                                {file && (<Button color="success" onPress={handleExtract}>
                                    Trích xuất
                                </Button>)}
                            </ModalFooter>
                        </>
                    )}
                </ModalContent>
            </Modal>

            <Modal
                size='5xl'
                isOpen={isPreviewOpen}
                onOpenChange={onPreviewOpenChange}
                isDismissable={false}
                isKeyboardDismissDisabled={true}
            >
                <ModalContent>
                    {(onClose) => (
                        <>
                            <ModalHeader className="flex flex-col gap-1">Dữ liệu trích xuất</ModalHeader>
                            <ModalBody>
                                <Table
                                    aria-label="Table with dynamic content"
                                    maxTableHeight={400}
                                    isVirtualized
                                >
                                    <TableHeader columns={[
                                        { key: 'stt', label: 'STT' },
                                        ...fields.map(f => ({ key: f.fieldName, label: f.nameDisplay }))
                                    ]}>
                                        {(column) => <TableColumn key={column.key}>{column.label}</TableColumn>}
                                    </TableHeader>
                                    <TableBody items={previewData}>
                                        {(item) => (
                                            <TableRow key={item.key}>
                                                {(columnKey) => (
                                                    <TableCell>{item[columnKey as keyof typeof item]}</TableCell>
                                                )}
                                            </TableRow>
                                        )}
                                    </TableBody>
                                </Table>
                            </ModalBody>
                            <ModalFooter>
                                <Button color="danger" onPress={onClose}>
                                    Đóng
                                </Button>
                                {cellError.length === 0 &&
                                    <Button color="primary" onPress={() => {
                                        console.log(extractedData);
                                        onClose();
                                        alert('Xuất thành công!')
                                    }}>
                                        Xuất
                                    </Button>}
                            </ModalFooter>
                        </>
                    )}
                </ModalContent>
            </Modal>
        </>
    );
}