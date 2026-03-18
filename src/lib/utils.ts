import { type ClassValue, clsx } from "clsx"
import { twMerge } from "tailwind-merge"
import { format } from 'date-fns';
import { ro } from 'date-fns/locale';
import XLSX from 'xlsx-js-style';
import { saveAs } from 'file-saver';
import JSZip from 'jszip';
import { Comanda, Doctor, Produs, Pacient } from "./types";

export function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs))
}

export const formatDate = (date: Date | string | undefined) => {
    if (!date) return 'N/A';
    if (typeof date === 'string') {
        // Plain date without time: extract DD/MM/YYYY directly
        const plainMatch = date.match(/^(\d{4})-(\d{2})-(\d{2})$/);
        if (plainMatch) return `${plainMatch[3]}/${plainMatch[2]}/${plainMatch[1]}`;
        // ISO timestamp: parse as Date to get local date (consistent with export filter)
        if (/^\d{4}-\d{2}-\d{2}[T ]/.test(date)) {
            const d = new Date(date);
            if (!isNaN(d.getTime())) {
                return format(d, 'dd/MM/yyyy', { locale: ro });
            }
        }
        // Handle DD/MM/YYYY, DD.MM.YYYY, DD-MM-YYYY
        const dmyMatch = date.match(/^(\d{1,2})[\/\.\-](\d{1,2})[\/\.\-](\d{4})$/);
        if (dmyMatch) return `${dmyMatch[1].padStart(2, '0')}/${dmyMatch[2].padStart(2, '0')}/${dmyMatch[3]}`;
    }
    const dateObj = typeof date === 'string' ? new Date(date) : date;
    return format(dateObj, 'dd/MM/yyyy', { locale: ro });
};

export const formatCurrency = (amount: number) => {
    return new Intl.NumberFormat('ro-RO', { style: 'currency', currency: 'RON' }).format(amount);
};

// Helper to build a styled worksheet for a doctor using the "Fisa Laborator" template.
// Returns a XLSX workbook ready to be written.
const buildDoctorWorkbook = (
    doctor: Doctor,
    comenziDoctor: Comanda[],
    produse: Produs[]
): XLSX.WorkBook => {
    // Build one section per ORDER (not grouped by patient).
    // Each order shows: patient name, products, order total.
    const orderSections: { pacientName: string; products: { name: string; cantitate: number; pret: number }[] }[] = [];

    comenziDoctor.forEach(comanda => {
        const pacient = doctor.pacienti.find(p => p.id === comanda.id_pacient);
        const pacientName = pacient?.nume || 'N/A';

        const products: { name: string; cantitate: number; pret: number }[] = [];
        if (comanda.produse.length === 0) {
            console.warn('[Export] Comanda', comanda.id, '(pacient:', pacientName, ') - array produse gol (0 produse încărcate)');
        }
        comanda.produse.forEach(cp => {
            const produs = produse.find(p => p.id === cp.id_produs);
            if (produs) {
                products.push({
                    name: produs.nume,
                    cantitate: cp.cantitate,
                    pret: produs.pret,
                });
            } else {
                console.warn('[Export] Comanda', comanda.id, '- produs cu id_produs', cp.id_produs, 'nu a fost găsit în lista de produse');
            }
        });

        // Always include the order, even if no products resolved
        orderSections.push({ pacientName, products });
    });

    // Build sheet data row by row
    const sheetData: (string | number | null)[][] = [];
    const merges: XLSX.Range[] = [];

    // Row 0-1: Title "Fisa Laborator" (merged A1:D2)
    sheetData.push(['Fisa Laborator', null, null, null]);
    sheetData.push([null, null, null, null]);
    merges.push({ s: { r: 0, c: 0 }, e: { r: 1, c: 3 } });

    // Row 2: Doctor name (merged A3:D3)
    sheetData.push([`Dr. ${doctor.nume}`, null, null, null]);
    merges.push({ s: { r: 2, c: 0 }, e: { r: 2, c: 3 } });

    // Row 3: Empty
    sheetData.push([null, null, null, null]);

    // Row 4: Headers
    sheetData.push(['PACIENT', 'PRODUS', 'BUCĂȚI', 'PREȚ']);

    let currentRow = 5;
    const orderTotalRows: number[] = [];

    for (const section of orderSections) {
        const startRow = currentRow;

        // Sort products alphabetically
        section.products.sort((a, b) => a.name.localeCompare(b.name, 'ro'));

        if (section.products.length === 0) {
            // Order with no products: show patient name with empty product/qty/price
            sheetData.push([section.pacientName, '-', 0, 0]);
            currentRow++;
        } else {
            section.products.forEach((product, idx) => {
                sheetData.push([
                    idx === 0 ? section.pacientName : null,
                    product.name,
                    product.cantitate,
                    product.pret * product.cantitate,
                ]);
                currentRow++;
            });

            // Merge patient name cells vertically if more than one product
            if (section.products.length > 1) {
                merges.push({ s: { r: startRow, c: 0 }, e: { r: currentRow - 1, c: 0 } });
            }
        }

        // Total per order row (A-C merged)
        const orderTotal = section.products.reduce((sum, p) => sum + p.pret * p.cantitate, 0);
        sheetData.push(['Total pacient', null, null, orderTotal]);
        merges.push({ s: { r: currentRow, c: 0 }, e: { r: currentRow, c: 2 } });
        orderTotalRows.push(currentRow);
        currentRow++;
    }

    // Empty row before grand total
    sheetData.push([null, null, null, null]);
    currentRow++;

    // Grand total row
    const grandTotal = orderTotalRows.reduce((sum, row) => sum + (sheetData[row][3] as number), 0);
    sheetData.push(['TOTAL', null, null, grandTotal]);
    const grandTotalRow = currentRow;

    // Create worksheet
    const ws = XLSX.utils.aoa_to_sheet(sheetData);
    ws['!merges'] = merges;

    // Column widths
    ws['!cols'] = [
        { wch: 25 }, // PACIENT
        { wch: 30 }, // PRODUS
        { wch: 12 }, // BUCĂȚI
        { wch: 15 }, // PREȚ
    ];

    // Row heights for title
    ws['!rows'] = [];
    ws['!rows'][0] = { hpt: 22 };
    ws['!rows'][1] = { hpt: 22 };

    // --- STYLES ---

    const allBorders = {
        top: { style: 'thin', color: { rgb: "000000" } },
        bottom: { style: 'thin', color: { rgb: "000000" } },
        left: { style: 'thin', color: { rgb: "000000" } },
        right: { style: 'thin', color: { rgb: "000000" } },
    };

    // Title: "Fisa Laborator" (A1:D2 merged)
    const cellA1 = XLSX.utils.encode_cell({ r: 0, c: 0 });
    ws[cellA1].s = {
        font: { name: 'Calibri', sz: 16, bold: true },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: allBorders,
    };
    // Ensure all merged title cells (B1, C1, D1, A2-D2) get borders
    for (let r = 0; r <= 1; r++) {
        for (let c = 0; c < 4; c++) {
            if (r === 0 && c === 0) continue; // already styled above
            const ref = XLSX.utils.encode_cell({ r, c });
            if (!ws[ref]) ws[ref] = { t: 's', v: '' };
            ws[ref].s = { border: allBorders };
        }
    }

    // Doctor name (A3:D3 merged)
    const cellA3 = XLSX.utils.encode_cell({ r: 2, c: 0 });
    if (ws[cellA3]) {
        ws[cellA3].s = {
            font: { name: 'Calibri', sz: 12, bold: true },
            fill: { fgColor: { rgb: "D3D3D3" } },
            alignment: { horizontal: 'center', vertical: 'center' },
            border: allBorders,
        };
    }
    // Ensure merged doctor name cells (B3, C3, D3) get borders
    for (let c = 1; c < 4; c++) {
        const ref = XLSX.utils.encode_cell({ r: 2, c });
        if (!ws[ref]) ws[ref] = { t: 's', v: '' };
        ws[ref].s = { fill: { fgColor: { rgb: "D3D3D3" } }, border: allBorders };
    }

    // Empty row 3 (row index 3) - apply borders
    for (let c = 0; c < 4; c++) {
        const ref = XLSX.utils.encode_cell({ r: 3, c });
        if (!ws[ref]) ws[ref] = { t: 's', v: '' };
        ws[ref].s = { border: allBorders };
    }

    // Header row (row 4)
    for (let c = 0; c < 4; c++) {
        const cellRef = XLSX.utils.encode_cell({ r: 4, c });
        if (ws[cellRef]) {
            ws[cellRef].s = {
                font: { name: 'Calibri', sz: 11, bold: true, color: { rgb: "274E13" } },
                fill: { fgColor: { rgb: "D9EAD3" } },
                alignment: { horizontal: 'center', vertical: 'center' },
                border: allBorders,
            };
        }
    }

    // Data cells (patient products)
    // Columns A (patient), B (product), C (quantity) = white background
    // Column D (price) = green background #E8F5E9
    const dataCellStyleWhite = {
        fill: { fgColor: { rgb: "FFFFFF" } },
        alignment: { horizontal: 'center', vertical: 'center' },
        border: allBorders,
    };
    const dataCellStyleGreen = {
        fill: { fgColor: { rgb: "E8F5E9" } },
        alignment: { horizontal: 'center', vertical: 'center' },
        numFmt: '0.00',
        border: allBorders,
    };

    for (let r = 5; r < sheetData.length; r++) {
        if (orderTotalRows.includes(r) || r === grandTotalRow) continue;
        // For empty rows (separators), just apply borders
        if (sheetData[r].every(v => v === null)) {
            for (let c = 0; c < 4; c++) {
                const cellRef = XLSX.utils.encode_cell({ r, c });
                if (!ws[cellRef]) ws[cellRef] = { t: 's', v: '' };
                ws[cellRef].s = { border: allBorders };
            }
            continue;
        }
        for (let c = 0; c < 4; c++) {
            const cellRef = XLSX.utils.encode_cell({ r, c });
            if (!ws[cellRef]) ws[cellRef] = { t: 's', v: '' };
            ws[cellRef].s = c < 3 ? dataCellStyleWhite : dataCellStyleGreen;
        }
    }

    // Order total rows - background #fef5e7, D text blue, label bold
    orderTotalRows.forEach(row => {
        const labelRef = XLSX.utils.encode_cell({ r: row, c: 0 });
        if (ws[labelRef]) {
            ws[labelRef].s = {
                font: { bold: true },
                fill: { fgColor: { rgb: "FEF5E7" } },
                alignment: { horizontal: 'center', vertical: 'center' },
                border: allBorders,
            };
        }
        // Ensure merged empty cells B and C also get background
        for (let c = 1; c <= 2; c++) {
            const ref = XLSX.utils.encode_cell({ r: row, c });
            if (!ws[ref]) ws[ref] = { t: 's', v: '' };
            ws[ref].s = {
                fill: { fgColor: { rgb: "FEF5E7" } },
                border: allBorders,
            };
        }
        const valueRef = XLSX.utils.encode_cell({ r: row, c: 3 });
        if (ws[valueRef]) {
            ws[valueRef].s = {
                fill: { fgColor: { rgb: "FEF5E7" } },
                font: { bold: true, color: { rgb: "0000FF" } },
                alignment: { horizontal: 'center', vertical: 'center' },
                numFmt: '0.00',
                border: allBorders,
            };
        }
    });

    // Grand total row - background #fff3cd, bold, A left-aligned
    for (let c = 0; c < 4; c++) {
        const cellRef = XLSX.utils.encode_cell({ r: grandTotalRow, c });
        if (!ws[cellRef]) ws[cellRef] = { t: 's', v: '' };
        ws[cellRef].s = {
            font: { bold: true },
            fill: { fgColor: { rgb: "FFF3CD" } },
            alignment: { horizontal: c === 0 ? 'left' : 'center', vertical: 'center' },
            border: allBorders,
            ...(c === 3 ? { numFmt: '0.00' } : {}),
        };
    }

    const wb = XLSX.utils.book_new();
    const safeSheetName = doctor.nume.replace(/[:\\/?*[\]]/g, '').substring(0, 31);
    XLSX.utils.book_append_sheet(wb, ws, safeSheetName);
    return wb;
};

// Convert a Date object to a local YYYY-MM-DD string.
const dateToLocalStr = (d: Date): string =>
    `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;

// Extract a YYYY-MM-DD date string from various date formats.
// Uses LOCAL timezone for ISO timestamps so the result matches the DatePicker dates.
const extractDateStr = (dateStr: string): string | null => {
    if (!dateStr) return null;
    const trimmed = String(dateStr).trim();

    // 1. Plain date without time: "2026-03-15" — use as-is (no timezone to convert)
    if (/^\d{4}-\d{2}-\d{2}$/.test(trimmed)) {
        return trimmed;
    }

    // 2. ISO / Supabase timestamp: "2026-03-15T14:30:00Z", "2026-03-15 14:30:00+00:00", etc.
    //    Parse via Date to get the LOCAL date (matching the user's timezone, same as DatePicker).
    if (/^\d{4}-\d{2}-\d{2}[T ]/.test(trimmed)) {
        const d = new Date(trimmed);
        if (!isNaN(d.getTime())) {
            return dateToLocalStr(d);
        }
        // Fallback: extract YYYY-MM-DD directly if Date parsing fails
        const m = trimmed.match(/^(\d{4})-(\d{2})-(\d{2})/);
        if (m) return `${m[1]}-${m[2]}-${m[3]}`;
    }

    // 3. DD/MM/YYYY or DD.MM.YYYY or DD-MM-YYYY (Romanian / European formats)
    const dmyMatch = trimmed.match(/^(\d{1,2})[\/\.\-](\d{1,2})[\/\.\-](\d{4})$/);
    if (dmyMatch) {
        return `${dmyMatch[3]}-${dmyMatch[2].padStart(2, '0')}-${dmyMatch[1].padStart(2, '0')}`;
    }

    // 4. Last resort: parse with Date constructor (uses local timezone)
    const d = new Date(trimmed);
    if (!isNaN(d.getTime())) {
        return dateToLocalStr(d);
    }

    return null;
};

// Check if status represents a finalized order (handles diacritics / casing / Unicode forms)
const isStatusFinalized = (status: string | undefined | null): boolean => {
    if (!status) return false;
    // Normalize to NFC first to handle decomposed Unicode characters (e.g., a + combining breve vs ă)
    const nfc = status.normalize('NFC');
    if (nfc === 'Finalizată') return true;
    const stripped = nfc.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
    return stripped === 'finalizata';
};

export const exportComenziToExcel = (
    comenzi: Comanda[],
    doctori: Doctor[],
    produse: Produs[],
    startDate: Date,
    endDate: Date
) => {
    const startStr = dateToLocalStr(startDate);
    const endStr = dateToLocalStr(endDate);

    console.debug('[Export] Date range:', startStr, '→', endStr);
    console.debug('[Export] Total comenzi:', comenzi.length,
        '| Finalizate:', comenzi.filter(c => isStatusFinalized(c.status)).length,
        '| Cu data_finalizare:', comenzi.filter(c => !!c.data_finalizare).length);

    // Log every finalized order for diagnostics
    comenzi.filter(c => isStatusFinalized(c.status)).forEach(c => {
        console.debug('[Export] Comanda finalizată', c.id,
            '| status:', JSON.stringify(c.status),
            '| data_finalizare:', JSON.stringify(c.data_finalizare),
            '| data_start:', JSON.stringify(c.data_start),
            '| termen_limita:', JSON.stringify(c.termen_limita));
    });

    const filteredComenzi = comenzi.filter(c => {
        if (!isStatusFinalized(c.status)) return false;

        // Filter by data_finalizare: include orders whose completion date
        // falls within the selected date range [startStr, endStr].
        const finDateStr = c.data_finalizare ? extractDateStr(c.data_finalizare) : null;

        if (!finDateStr) {
            console.debug('[Export] Comanda', c.id, '- fără data_finalizare, exclusă');
            return false;
        }

        const inRange = finDateStr >= startStr && finDateStr <= endStr;
        if (!inRange) {
            console.debug('[Export] Comanda', c.id, '- data_finalizare', finDateStr, 'nu este în intervalul', startStr, '-', endStr);
        } else {
            console.debug('[Export] Comanda', c.id, '- INCLUSĂ, data_finalizare:', finDateStr);
        }
        return inRange;
    });

    console.debug('[Export] Comenzi filtrate:', filteredComenzi.length);
    // Log product counts for each filtered order
    filteredComenzi.forEach(c => {
        console.debug('[Export] Comanda', c.id, '| produse:', c.produse.length, '| produse ids:', c.produse.map(p => p.id_produs));
    });

    const groupedByDoctor = filteredComenzi.reduce((acc, comanda) => {
        (acc[comanda.id_doctor] = acc[comanda.id_doctor] || []).push(comanda);
        return acc;
    }, {} as Record<number, Comanda[]>);

    for (const doctorId in groupedByDoctor) {
        const doctor = doctori.find(d => d.id === Number(doctorId));
        if (!doctor) continue;

        const wb = buildDoctorWorkbook(doctor, groupedByDoctor[doctorId], produse);
        const filename = `${doctor.nume.replace(/\s/g, '_')}_${format(startDate, 'dd-MM-yyyy')}_${format(endDate, 'dd-MM-yyyy')}.xlsx`;
        XLSX.writeFile(wb, filename);
    }
};

export const exportAllComenziToZip = async (
    comenzi: Comanda[],
    doctori: Doctor[],
    produse: Produs[],
    startDate: Date,
    endDate: Date
) => {
    const startStr = dateToLocalStr(startDate);
    const endStr = dateToLocalStr(endDate);

    console.debug('[ExportZip] Date range:', startStr, '→', endStr);
    console.debug('[ExportZip] Total comenzi:', comenzi.length,
        '| Finalizate:', comenzi.filter(c => isStatusFinalized(c.status)).length,
        '| Cu data_finalizare:', comenzi.filter(c => !!c.data_finalizare).length);

    // Log every finalized order for diagnostics
    comenzi.filter(c => isStatusFinalized(c.status)).forEach(c => {
        console.debug('[ExportZip] Comanda finalizată', c.id,
            '| status:', JSON.stringify(c.status),
            '| data_finalizare:', JSON.stringify(c.data_finalizare),
            '| data_start:', JSON.stringify(c.data_start),
            '| termen_limita:', JSON.stringify(c.termen_limita));
    });

    const filteredComenzi = comenzi.filter(c => {
        if (!isStatusFinalized(c.status)) return false;

        // Filter by data_finalizare: include orders whose completion date
        // falls within the selected date range [startStr, endStr].
        const finDateStr = c.data_finalizare ? extractDateStr(c.data_finalizare) : null;

        if (!finDateStr) {
            console.debug('[ExportZip] Comanda', c.id, '- fără data_finalizare, exclusă');
            return false;
        }

        const inRange = finDateStr >= startStr && finDateStr <= endStr;
        if (!inRange) {
            console.debug('[ExportZip] Comanda', c.id, '- data_finalizare', finDateStr, 'nu este în intervalul', startStr, '-', endStr);
        } else {
            console.debug('[ExportZip] Comanda', c.id, '- INCLUSĂ, data_finalizare:', finDateStr);
        }
        return inRange;
    });

    console.debug('[ExportZip] Comenzi filtrate:', filteredComenzi.length);
    // Log product counts for each filtered order
    filteredComenzi.forEach(c => {
        console.debug('[ExportZip] Comanda', c.id, '| produse:', c.produse.length, '| produse ids:', c.produse.map(p => p.id_produs));
    });

    if (filteredComenzi.length === 0) {
        throw new Error('Nu există comenzi cu data de finalizare în perioada selectată.');
    }

    const groupedByDoctor = filteredComenzi.reduce((acc, comanda) => {
        (acc[comanda.id_doctor] = acc[comanda.id_doctor] || []).push(comanda);
        return acc;
    }, {} as Record<number, Comanda[]>);

    const zip = new JSZip();

    for (const doctorId in groupedByDoctor) {
        const doctor = doctori.find(d => d.id === Number(doctorId));
        if (!doctor) continue;

        const wb = buildDoctorWorkbook(doctor, groupedByDoctor[doctorId], produse);
        const wbBinary = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        const filename = `${doctor.nume.replace(/\s/g, '_')}.xlsx`;
        zip.file(filename, wbBinary);
    }

    const zipBlob = await zip.generateAsync({ type: 'blob' });
    const startFmt = format(startDate, 'dd-MM-yyyy');
    const endFmt = format(endDate, 'dd-MM-yyyy');
    saveAs(zipBlob, `Export_Comenzi_${startFmt}_${endFmt}.zip`);
};

export const generateOrderExcel = (
    comanda: Comanda,
    doctor: Doctor | undefined,
    pacient: Pacient | undefined,
    produse: Produs[]
) => {
    const sheetData: (string | number | null)[][] = [];
    const merges: XLSX.Range[] = [];

    // Row 0 (Line 1): "Fișă de laborator" - Merge A1:C1
    sheetData.push(['Fișă de laborator', null, null]);
    merges.push({ s: { r: 0, c: 0 }, e: { r: 0, c: 2 } });

    // Row 1 (Line 2): Doctor - Merge B2:C2
    sheetData.push(['Nume Doctor:', doctor?.nume || 'N/A', null]);
    merges.push({ s: { r: 1, c: 1 }, e: { r: 1, c: 2 } });

    // Row 2 (Line 3): Patient - Merge B3:C3
    sheetData.push(['Nume Pacient:', pacient?.nume || 'N/A', null]);
    merges.push({ s: { r: 2, c: 1 }, e: { r: 2, c: 2 } });

    // Row 3 (Line 4): "Produse" - Merge A4:C4
    sheetData.push(['Produse', null, null]);
    merges.push({ s: { r: 3, c: 0 }, e: { r: 3, c: 2 } });

    // Starting from Row 4 (Line 5): Products
    let grandTotal = 0;
    comanda.produse.forEach(cp => {
        const produs = produse.find(p => p.id === cp.id_produs);
        const productTotal = (produs?.pret || 0) * cp.cantitate;
        grandTotal += productTotal;
        sheetData.push([
            produs?.nume || 'N/A',
            `${cp.cantitate} buc`,
            productTotal,
        ]);
    });

    // Last line: "Total: {total general}" - Merge A:C
    const lastRowIndex = sheetData.length;
    sheetData.push([`Total: ${grandTotal.toFixed(2)}`, null, null]);
    merges.push({ s: { r: lastRowIndex, c: 0 }, e: { r: lastRowIndex, c: 2 } });

    // Create worksheet
    const ws = XLSX.utils.aoa_to_sheet(sheetData);
    ws['!merges'] = merges;

    // Column widths
    ws['!cols'] = [
        { wch: 30 },
        { wch: 15 },
        { wch: 15 },
    ];

    // --- STYLES ---
    const allBorders = {
        top: { style: 'thin', color: { rgb: "000000" } },
        bottom: { style: 'thin', color: { rgb: "000000" } },
        left: { style: 'thin', color: { rgb: "000000" } },
        right: { style: 'thin', color: { rgb: "000000" } },
    };

    const orangeFill = { fgColor: { rgb: "EBF1DE" } };
    const lightOrangeFill = { fgColor: { rgb: "F4B084" } };
    const paleOrangeFill = { fgColor: { rgb: "FCE4D6" } };

    // Helper to ensure a cell exists and apply style
    const setStyle = (r: number, c: number, style: any) => {
        const ref = XLSX.utils.encode_cell({ r, c });
        if (!ws[ref]) ws[ref] = { t: 's', v: '' };
        ws[ref].s = style;
    };

    // Line 1 (Row 0): "Fișă de laborator" - #ebf1de
    for (let c = 0; c < 3; c++) {
        setStyle(0, c, {
            font: { name: 'Calibri', sz: 14, bold: true, color: { rgb: "000000" } },
            fill: orangeFill,
            alignment: { horizontal: 'center', vertical: 'center' },
            border: allBorders,
        });
    }

    // Lines 2 and 3 (Rows 1-2): Doctor and Patient - #fce4d6
    for (let r = 1; r <= 2; r++) {
        setStyle(r, 0, {
            font: { name: 'Calibri', bold: true },
            fill: paleOrangeFill,
            alignment: { vertical: 'center' },
            border: allBorders,
        });
        for (let c = 1; c <= 2; c++) {
            setStyle(r, c, {
                font: { name: 'Calibri' },
                fill: paleOrangeFill,
                alignment: { vertical: 'center' },
                border: allBorders,
            });
        }
    }

    // Line 4 (Row 3): "Produse" - #ebf1de
    for (let c = 0; c < 3; c++) {
        setStyle(3, c, {
            font: { name: 'Calibri', sz: 12, bold: true, color: { rgb: "000000" } },
            fill: orangeFill,
            alignment: { horizontal: 'center', vertical: 'center' },
            border: allBorders,
        });
    }

    // Product rows (Row 4 to lastRowIndex - 1)
    for (let r = 4; r < lastRowIndex; r++) {
        // Column A: product name - #fce4d6
        setStyle(r, 0, {
            font: { name: 'Calibri' },
            fill: paleOrangeFill,
            alignment: { vertical: 'center' },
            border: allBorders,
        });
        // Column B: quantity - #fce4d6
        setStyle(r, 1, {
            font: { name: 'Calibri' },
            fill: paleOrangeFill,
            alignment: { horizontal: 'center', vertical: 'center' },
            border: allBorders,
        });
        // Column C: total product - #f4b084
        setStyle(r, 2, {
            font: { name: 'Calibri' },
            fill: lightOrangeFill,
            alignment: { horizontal: 'center', vertical: 'center' },
            border: allBorders,
            numFmt: '0.00',
        });
    }

    // Last line (Row lastRowIndex): "Total: {total}" - #ebf1de
    for (let c = 0; c < 3; c++) {
        setStyle(lastRowIndex, c, {
            font: { name: 'Calibri', sz: 12, bold: true, color: { rgb: "000000" } },
            fill: orangeFill,
            alignment: { horizontal: 'center', vertical: 'center' },
            border: allBorders,
        });
    }

    // Create workbook and save
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Fisa de laborator');

    const filename = `Fisa_Laborator_${doctor?.nume?.replace(/\s/g, '_') || 'N-A'}_${pacient?.nume?.replace(/\s/g, '_') || 'N-A'}_${comanda.id}.xlsx`;
    XLSX.writeFile(wb, filename);
};
