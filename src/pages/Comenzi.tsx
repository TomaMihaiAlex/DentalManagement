import React, { useState, useMemo } from 'react';
import { useData } from '@/context/DataContext';
import { Button } from '@/components/ui/Button';
import { Card, CardContent, CardHeader, CardTitle, CardFooter } from '@/components/ui/Card';
import { Input } from '@/components/ui/Input';
import { PlusCircle, Edit, Trash2, Search, Check, Play, Clock, FileDown, RefreshCcw, Banknote, HardHat, CalendarCheck, XCircle, Printer, ChevronLeft, ChevronRight } from 'lucide-react';
import { Comanda, OrderStatus } from '@/lib/types';
import { formatDate, formatCurrency, exportAllComenziToZip, generateOrderWordDocument } from '@/lib/utils';
import ConfirmDeleteModal from '@/components/modals/ConfirmDeleteModal';
import ComandaModal from '@/components/modals/ComandaModal';
import FinalizeazaComandaModal from '@/components/modals/FinalizeazaComandaModal';
import ConfirmFinalizeModal from '@/components/modals/ConfirmFinalizeModal';
import ExportExcelModal from '@/components/modals/ExportExcelModal';
import toast from 'react-hot-toast';
import OrderStatusProgress from '@/components/OrderStatusProgress';

const statusConfig: Record<OrderStatus, { text: string; bg: string; icon: React.ElementType }> = {
    'Finalizată': { text: 'text-success', bg: 'bg-green-100 dark:bg-green-900/50', icon: Check },
    'În progres': { text: 'text-warning', bg: 'bg-orange-100 dark:bg-orange-900/50', icon: Play },
    'Întârziată': { text: 'text-danger', bg: 'bg-red-100 dark:bg-red-900/50', icon: Clock },
};

const Comenzi: React.FC = () => {
    const { comenzi, doctori, produse, addComanda, updateComanda, deleteComanda, finalizeComanda, reopenComanda } = useData();
    const [statusFilter, setStatusFilter] = useState<OrderStatus | 'Toate'>('Toate');
    
    // Search states
    const [doctorSearch, setDoctorSearch] = useState('');
    const [pacientSearch, setPacientSearch] = useState('');
    const [tehnicianSearch, setTehnicianSearch] = useState('');
    const [appliedFilters, setAppliedFilters] = useState({ doctor: '', pacient: '', tehnician: '' });

    // Modal States
    const [isComandaModalOpen, setComandaModalOpen] = useState(false);
    const [isDeleteModalOpen, setDeleteModalOpen] = useState(false);
    const [isReopenModalOpen, setReopenModalOpen] = useState(false);
    const [isFinalizeModalOpen, setFinalizeModalOpen] = useState(false);
    const [isConfirmFinalizeModalOpen, setConfirmFinalizeModalOpen] = useState(false);
    const [isExportModalOpen, setExportModalOpen] = useState(false);

    // Data for Modals
    const [selectedComanda, setSelectedComanda] = useState<Comanda | null>(null);
    const [dataForFinalization, setDataForFinalization] = useState<{ comanda: Comanda, tehnician: string } | null>(null);

    // Pagination
    const ITEMS_PER_PAGE = 20;
    const [currentPage, setCurrentPage] = useState(1);

    const handleSearch = () => {
        setAppliedFilters({
            doctor: doctorSearch,
            pacient: pacientSearch,
            tehnician: tehnicianSearch,
        });
        setCurrentPage(1);
    };

    const handleResetSearch = () => {
        setDoctorSearch('');
        setPacientSearch('');
        setTehnicianSearch('');
        setAppliedFilters({ doctor: '', pacient: '', tehnician: '' });
        setCurrentPage(1);
    };

    const filteredComenzi = useMemo(() => {
        return comenzi
            .filter(c => statusFilter === 'Toate' || c.status === statusFilter)
            .filter(c => {
                const doctor = doctori.find(d => d.id === c.id_doctor);
                const pacient = doctor?.pacienti.find(p => p.id === c.id_pacient);

                const doctorMatch = !appliedFilters.doctor || doctor?.nume.toLowerCase().includes(appliedFilters.doctor.toLowerCase());
                const pacientMatch = !appliedFilters.pacient || pacient?.nume.toLowerCase().includes(appliedFilters.pacient.toLowerCase());
                const tehnicianMatch = !appliedFilters.tehnician || (c.tehnician && c.tehnician.toLowerCase().includes(appliedFilters.tehnician.toLowerCase()));

                return doctorMatch && pacientMatch && tehnicianMatch;
            })
            .sort((a, b) => b.id - a.id);
    }, [comenzi, statusFilter, appliedFilters, doctori]);
    
    const progressStats = useMemo(() => {
        const total = comenzi.length;
        const completed = comenzi.filter(c => c.status === 'Finalizată').length;
        const inProgress = comenzi.filter(c => c.status === 'În progres').length;
        const delayed = comenzi.filter(c => c.status === 'Întârziată').length;
        return { completed, inProgress, delayed, total };
    }, [comenzi]);

    // Pagination computed values
    const totalPages = Math.max(1, Math.ceil(filteredComenzi.length / ITEMS_PER_PAGE));
    const paginatedComenzi = useMemo(() => {
        const start = (currentPage - 1) * ITEMS_PER_PAGE;
        return filteredComenzi.slice(start, start + ITEMS_PER_PAGE);
    }, [filteredComenzi, currentPage, ITEMS_PER_PAGE]);

    // Reset page when status filter changes
    const handleStatusFilter = (f: OrderStatus | 'Toate') => {
        setStatusFilter(f);
        setCurrentPage(1);
    };

    // Generate page numbers for pagination
    const getPageNumbers = () => {
        const pages: (number | '...')[] = [];
        if (totalPages <= 7) {
            for (let i = 1; i <= totalPages; i++) pages.push(i);
        } else {
            pages.push(1);
            if (currentPage > 3) pages.push('...');
            for (let i = Math.max(2, currentPage - 1); i <= Math.min(totalPages - 1, currentPage + 1); i++) {
                pages.push(i);
            }
            if (currentPage < totalPages - 2) pages.push('...');
            pages.push(totalPages);
        }
        return pages;
    };

    const handleSaveComanda = (comandaData: any) => {
        if (comandaData.id) {
            updateComanda(comandaData);
        } else {
            addComanda(comandaData);
        }
    };

    const handleDeleteConfirm = () => {
        if(selectedComanda) {
            deleteComanda(selectedComanda.id);
            setDeleteModalOpen(false);
            setSelectedComanda(null);
        }
    };

    const handleOpenReopenModal = (comanda: Comanda) => {
        setSelectedComanda(comanda);
        setReopenModalOpen(true);
    };

    const handleReopenConfirm = () => {
        if (selectedComanda) {
            reopenComanda(selectedComanda.id);
            setReopenModalOpen(false);
            setSelectedComanda(null);
        }
    };

    const handleProceedToFinalize = (tehnician: string) => {
        if (selectedComanda) {
            setDataForFinalization({ comanda: selectedComanda, tehnician });
            setFinalizeModalOpen(false);
            setConfirmFinalizeModalOpen(true);
        }
    };

    const handleFinalizeConfirm = () => {
        if (dataForFinalization) {
            finalizeComanda(dataForFinalization.comanda.id, dataForFinalization.tehnician);
            setConfirmFinalizeModalOpen(false);
            setDataForFinalization(null);
            setSelectedComanda(null);
        }
    };

    const handleExport = async (startDate: Date, endDate: Date) => {
        try {
            await exportAllComenziToZip(comenzi, doctori, produse, startDate, endDate);
            toast.success("Arhiva ZIP a fost descărcată cu succes.");
        } catch (error: any) {
            console.error("Export failed:", error);
            toast.error(error?.message || "Exportul a eșuat.");
        }
    };

    const handlePrintOrder = async (comanda: Comanda) => {
        try {
            const doctor = doctori.find(d => d.id === comanda.id_doctor);
            const pacient = doctor?.pacienti.find(p => p.id === comanda.id_pacient);
            await generateOrderWordDocument(comanda, doctor, pacient, produse);
            toast.success("Documentul a fost generat cu succes.");
        } catch (error) {
            console.error("Print failed:", error);
            toast.error("Generarea documentului a eșuat.");
        }
    };

    return (
        <div className="space-y-6">
            <Card>
                <CardHeader>
                    <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-4">
                        <CardTitle>Management Comenzi</CardTitle>
                        <div className="flex items-center gap-2">
                             <Button variant="outline" onClick={() => setExportModalOpen(true)}>
                                <FileDown className="w-4 h-4 mr-2" />
                                Export
                            </Button>
                            <Button onClick={() => { setSelectedComanda(null); setComandaModalOpen(true); }}>
                                <PlusCircle className="w-4 h-4 mr-2" />
                                Adaugă
                            </Button>
                        </div>
                    </div>
                    <div className="mt-4 p-4 border rounded-lg dark:border-gray-700">
                        <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                            <Input placeholder="Caută după doctor..." value={doctorSearch} onChange={e => setDoctorSearch(e.target.value)} />
                            <Input placeholder="Caută după pacient..." value={pacientSearch} onChange={e => setPacientSearch(e.target.value)} />
                            <Input placeholder="Caută după tehnician..." value={tehnicianSearch} onChange={e => setTehnicianSearch(e.target.value)} />
                        </div>
                        <div className="flex justify-end gap-2 mt-4">
                            <Button variant="ghost" onClick={handleResetSearch}>
                                <XCircle className="w-4 h-4 mr-2"/>
                                Reset
                            </Button>
                            <Button onClick={handleSearch}>
                                <Search className="w-4 h-4 mr-2"/>
                                Caută
                            </Button>
                        </div>
                    </div>
                    <div className="flex flex-wrap gap-2 mt-4">
                        {(['Toate', 'În progres', 'Finalizată', 'Întârziată'] as const).map(f => (
                            <Button key={f} variant={statusFilter === f ? 'default' : 'secondary'} size="sm" onClick={() => handleStatusFilter(f)}>{f}</Button>))}
                    </div>
                </CardHeader>
                <CardContent>
                    <OrderStatusProgress stats={progressStats} />
                </CardContent>
            </Card>

            {/* Mobile & Tablet Card View */}
            <div className="grid grid-cols-1 lg:hidden gap-4">
                {paginatedComenzi.map(comanda => {
                    const doctor = doctori.find(d => d.id === comanda.id_doctor);
                    const pacient = doctor?.pacienti.find(p => p.id === comanda.id_pacient);
                    const status = statusConfig[comanda.status];
                    const isFinalized = comanda.status === 'Finalizată';
                    return (
                        <Card key={comanda.id} className="flex flex-col">
                            <CardHeader>
                                <div className="flex justify-between items-start">
                                    <div>
                                        <CardTitle className="text-base">{doctor?.nume}</CardTitle>
                                        <p className="text-sm text-gray-500 dark:text-gray-400">{pacient?.nume}</p>
                                    </div>
                                    <span className={`inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium ${status.bg} ${status.text}`}>
                                        <status.icon className="w-3 h-3 mr-1.5" />
                                        {comanda.status}
                                    </span>
                                </div>
                            </CardHeader>
                            <CardContent className="space-y-2 text-sm flex-grow">
                                <div className="flex items-center"><Banknote className="w-4 h-4 mr-2 text-gray-500"/> <strong>Total:</strong><span className="ml-auto">{formatCurrency(comanda.total)}</span></div>
                                <div className="flex items-center"><CalendarCheck className="w-4 h-4 mr-2 text-gray-500"/> <strong>Start:</strong><span className="ml-auto">{formatDate(comanda.data_start)}</span></div>
                                <div className="flex items-center"><CalendarCheck className="w-4 h-4 mr-2 text-gray-500"/> <strong>Termen:</strong><span className="ml-auto">{formatDate(comanda.termen_limita)}</span></div>
                                {comanda.data_finalizare && <div className="flex items-center"><CalendarCheck className="w-4 h-4 mr-2 text-success"/> <strong>Finalizat:</strong><span className="ml-auto">{formatDate(comanda.data_finalizare)}</span></div>}
                                {comanda.tehnician && <div className="flex items-center"><HardHat className="w-4 h-4 mr-2 text-gray-500"/> <strong>Tehnician:</strong><span className="ml-auto">{comanda.tehnician}</span></div>}
                            </CardContent>
                            <CardFooter className="flex justify-end gap-2">
                                {isFinalized ? (
                                    <Button variant="outline" size="sm" onClick={() => handleOpenReopenModal(comanda)}><RefreshCcw className="w-4 h-4 mr-2"/>Reia</Button>
                                ) : (
                                    <Button variant="outline" size="sm" onClick={() => { setSelectedComanda(comanda); setFinalizeModalOpen(true); }}>Finalizează</Button>
                                )}
                                <Button variant="outline" size="sm" onClick={() => handlePrintOrder(comanda)}><Printer className="w-4 h-4 mr-2"/>Imprimare</Button>
                                <Button variant="ghost" size="icon" onClick={() => { setSelectedComanda(comanda); setComandaModalOpen(true); }}><Edit className="w-4 h-4" /></Button>
                                <Button variant="ghost" size="icon" onClick={() => { setSelectedComanda(comanda); setDeleteModalOpen(true); }}><Trash2 className="w-4 h-4 text-danger" /></Button>
                            </CardFooter>
                        </Card>
                    )
                })}
            </div>

            {/* Desktop Table View */}
            <div className="hidden lg:block">
                <Card>
                    <CardContent className="p-0">
                        <div className="overflow-x-auto">
                            <table className="w-full text-sm text-left text-gray-500 dark:text-gray-400">
                                <thead className="text-xs text-gray-700 uppercase bg-gray-50 dark:bg-gray-700 dark:text-gray-400">
                                    <tr>
                                        {['Doctor', 'Pacient', 'Total', 'Data Start', 'Termen Limită', 'Data Finalizare', 'Status', 'Tehnician', 'Acțiuni'].map(h => <th key={h} scope="col" className="px-6 py-3 whitespace-nowrap">{h}</th>)}
                                    </tr>
                                </thead>
                                <tbody>
                                    {paginatedComenzi.map(comanda => {
                                        const doctor = doctori.find(d => d.id === comanda.id_doctor);
                                        const pacient = doctor?.pacienti.find(p => p.id === comanda.id_pacient);
                                        const status = statusConfig[comanda.status];
                                        const isFinalized = comanda.status === 'Finalizată';
                                        return (
                                            <tr key={comanda.id} className="bg-white border-b dark:bg-gray-800 dark:border-gray-700 hover:bg-gray-50 dark:hover:bg-gray-600">
                                                <td className="px-6 py-4">{doctor?.nume || 'N/A'}</td>
                                                <td className="px-6 py-4">{pacient?.nume || 'N/A'}</td>
                                                <td className="px-6 py-4">{formatCurrency(comanda.total)}</td>
                                                <td className="px-6 py-4">{formatDate(comanda.data_start)}</td>
                                                <td className="px-6 py-4">{formatDate(comanda.termen_limita)}</td>
                                                <td className="px-6 py-4">{formatDate(comanda.data_finalizare)}</td>
                                                <td className="px-6 py-4">
                                                    <span className={`inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium ${status.bg} ${status.text}`}>
                                                        <status.icon className="w-3 h-3 mr-1.5" />
                                                        {comanda.status}
                                                    </span>
                                                </td>
                                                <td className="px-6 py-4">{comanda.tehnician || 'N/A'}</td>
                                                <td className="px-6 py-4 flex items-center gap-1">
                                                    {isFinalized ? (
                                                        <Button variant="outline" size="sm" onClick={() => handleOpenReopenModal(comanda)}><RefreshCcw className="w-4 h-4 mr-2"/>Reia</Button>
                                                    ) : (
                                                        <Button variant="outline" size="sm" onClick={() => { setSelectedComanda(comanda); setFinalizeModalOpen(true); }}>Finalizează</Button>
                                                    )}
                                                    <Button variant="outline" size="sm" onClick={() => handlePrintOrder(comanda)}><Printer className="w-4 h-4 mr-2"/>Imprimare</Button>
                                                    <Button variant="ghost" size="icon" onClick={() => { setSelectedComanda(comanda); setComandaModalOpen(true); }}><Edit className="w-4 h-4" /></Button>
                                                    <Button variant="ghost" size="icon" onClick={() => { setSelectedComanda(comanda); setDeleteModalOpen(true); }}><Trash2 className="w-4 h-4 text-danger" /></Button>
                                                </td>
                                            </tr>
                                        );
                                    })}
                                </tbody>
                            </table>
                        </div>
                    </CardContent>
                </Card>
            </div>

            {/* Pagination */}
            {totalPages > 1 && (
                <div className="flex flex-col sm:flex-row items-center justify-between gap-4">
                    <p className="text-sm text-gray-500 dark:text-gray-400">
                        Afișare {((currentPage - 1) * ITEMS_PER_PAGE) + 1}–{Math.min(currentPage * ITEMS_PER_PAGE, filteredComenzi.length)} din {filteredComenzi.length} comenzi
                    </p>
                    <div className="flex items-center gap-1">
                        <Button
                            variant="outline"
                            size="sm"
                            onClick={() => setCurrentPage(p => Math.max(1, p - 1))}
                            disabled={currentPage === 1}
                        >
                            <ChevronLeft className="w-4 h-4" />
                        </Button>
                        {getPageNumbers().map((page, idx) =>
                            page === '...' ? (
                                <span key={`dots-${idx}`} className="px-2 text-gray-400">...</span>
                            ) : (
                                <Button
                                    key={page}
                                    variant={currentPage === page ? 'default' : 'outline'}
                                    size="sm"
                                    onClick={() => setCurrentPage(page)}
                                    className="min-w-[36px] text-white"
                                >
                                    {page}
                                </Button>
                            )
                        )}
                        <Button
                            variant="outline"
                            size="sm"
                            onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))}
                            disabled={currentPage === totalPages}
                        >
                            <ChevronRight className="w-4 h-4" />
                        </Button>
                    </div>
                </div>
            )}

            <ComandaModal isOpen={isComandaModalOpen} onClose={() => setComandaModalOpen(false)} onSave={handleSaveComanda} comanda={selectedComanda} />
            <FinalizeazaComandaModal isOpen={isFinalizeModalOpen} onClose={() => setFinalizeModalOpen(false)} onConfirm={handleProceedToFinalize} />
            <ConfirmFinalizeModal isOpen={isConfirmFinalizeModalOpen} onClose={() => setConfirmFinalizeModalOpen(false)} onConfirm={handleFinalizeConfirm} data={dataForFinalization} />
            <ExportExcelModal isOpen={isExportModalOpen} onClose={() => setExportModalOpen(false)} onExport={handleExport} />
            <ConfirmDeleteModal isOpen={isDeleteModalOpen} onClose={() => setDeleteModalOpen(false)} onConfirm={handleDeleteConfirm} title="Confirmă Ștergerea Comenzii" description="Ești sigur că vrei să ștergi această comandă?" />
            <ConfirmDeleteModal isOpen={isReopenModalOpen} onClose={() => setReopenModalOpen(false)} onConfirm={handleReopenConfirm} title="Confirmă Redeschiderea" description="Ești sigur că vrei să redeschizi această comandă? Statutul va fi schimbat înapoi la 'În progres'." confirmText="Da, redeschide" />
        </div>
    );
};

export default Comenzi;
