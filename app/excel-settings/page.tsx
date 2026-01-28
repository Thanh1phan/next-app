'use client'

import React, { useState, useEffect } from 'react';
import { Plus, RefreshCw, EyeIcon, DeleteIcon } from 'lucide-react';
import { addToast, Button, ButtonGroup, Table, TableBody, TableCell, TableColumn, TableHeader, TableRow } from '@heroui/react';
import { API_BASE_URL, ConfigType, DepartmentId, ExcelConfig } from '@/types/excel-type';
import { useRouter } from 'next/navigation';

const ExcelConfigManager = () => {
    const [configs, setConfigs] = useState<ExcelConfig[]>([]);
    const [selectedConfig, setSelectedConfig] = useState<ExcelConfig | null>(null);
    const [isLoadingConfigs, setIsLoadingConfigs] = useState(true);
    const [error, setError] = useState<string | null>(null);
    const router = useRouter();

    // Fetch all configs
    const fetchConfigs = async () => {
        setIsLoadingConfigs(true);
        setError(null);
        try {
            const response = await fetch(`${API_BASE_URL}/excel-config`);
            if (!response.ok) throw new Error('Failed to fetch configs');
            const data = await response.json();
            setConfigs(data);
        } catch (err) {
            setError(err instanceof Error ? err.message : 'An error occurred');
            console.error('Error fetching configs:', err);
        } finally {
            setIsLoadingConfigs(false);
        }
    };

    // Initial load
    useEffect(() => {
        fetchConfigs();
    }, []);

    const handleRefreshConfigs = () => {
        fetchConfigs();
        setSelectedConfig(null);
    };

    const handleDeleteConfig = (detailId: number) => {
        if (!selectedConfig) return;

        const updatedConfigs = configs.map(config => {
            if (config.id === selectedConfig.id) {
                return {
                    ...config,
                    details: (config.details || []).filter(d => d.id !== detailId)
                };
            }
            return config;
        });

        setConfigs(updatedConfigs);
        setSelectedConfig(updatedConfigs.find(c => c.id === selectedConfig.id) || null);
    };

    const renderCell = (item: ExcelConfig, columnKey: React.Key) => {
        switch (columnKey) {
            case "index":
                return configs.indexOf(item) + 1;

            case "configName":
                return item.configName;

            case "departmentId":
                return DepartmentId[item.departmentId];

            case "configType":
                return ConfigType[item.configType];
            case "actions":
                return (
                    <ButtonGroup size='sm'>
                        <Button onPress={() => { router.push(`/excel-edit/${item.id}`); }}>
                            <EyeIcon />
                        </Button>
                        <Button
                            color="danger"
                            onPress={() => {
                                addToast({
                                    title: "Toast title",
                                    description: "Toast displayed successfully",
                                    radius: 'sm',
                                })
                            }}
                        >
                            <DeleteIcon />
                        </Button>
                    </ButtonGroup>
                );
            default:
                return null;
        }
    };


    return (
        <div className="min-h-screen p-6">
            <div className="mb-6">
                <div className="flex justify-between items-center">
                    <div>
                        <h1 className="text-3xl font-bold">Cấu hình Extract Excel</h1>
                        <p className="mt-2">Quản lý cấu hình import/export Excel</p>
                    </div>
                    <div className='grid grid-cols-2 gap-1'>
                        <Button
                            onPress={() => router.push('/excel-create')}
                            disabled={isLoadingConfigs}
                            color='success'
                        >
                            <Plus className='w-4 h-4' />
                        </Button>
                        <Button
                            onPress={fetchConfigs}
                            disabled={isLoadingConfigs}
                        >
                            <RefreshCw className={`w-4 h-4 ${isLoadingConfigs ? 'animate-spin' : ''}`} />
                        </Button>
                    </div>
                </div>
            </div>

            <Table
                aria-label="Table with dynamic content"
                maxTableHeight={500}
                isVirtualized
            >
                <TableHeader className='sticky top-0 '>
                    <TableColumn key="index">STT</TableColumn>
                    <TableColumn key="configName">Tên config</TableColumn>
                    <TableColumn key="departmentId">DepartmentId</TableColumn>
                    <TableColumn key="configType">Loại</TableColumn>
                    <TableColumn key="actions">Thao tác</TableColumn>
                </TableHeader>
                <TableBody items={configs}>
                    {(item) => (
                        <TableRow key={item.id}>
                            {(columnKey) => <TableCell>{renderCell(item, columnKey)}</TableCell>}
                        </TableRow>
                    )}
                </TableBody>
            </Table>

        </div >
    );
};

export default ExcelConfigManager;