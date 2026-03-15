import { type ClassValue, clsx } from "clsx"
import { twMerge } from "tailwind-merge"
import { format } from 'date-fns';
import { ro } from 'date-fns/locale';
import XLSX from 'xlsx-js-style';
import { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel } from 'docx';
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

    // Title: "Fisa Laborator" (A1)
    const cellA1 = XLSX.utils.encode_cell({ r: 0, c: 0 });
    ws[cellA1].s = {
        font: { name: 'Calibri', sz: 16, bold: true },
        alignment: { horizontal: 'center', vertical: 'center' },
    };

    // Doctor name (A3)
    const cellA3 = XLSX.utils.encode_cell({ r: 2, c: 0 });
    if (ws[cellA3]) {
        ws[cellA3].s = {
            font: { name: 'Calibri', sz: 12, bold: true },
            fill: { fgColor: { rgb: "D3D3D3" } },
            alignment: { horizontal: 'center', vertical: 'center' },
        };
    }

    // Header row (row 4)
    for (let c = 0; c < 4; c++) {
        const cellRef = XLSX.utils.encode_cell({ r: 4, c });
        if (ws[cellRef]) {
            ws[cellRef].s = {
                font: { name: 'Calibri', sz: 11, bold: true, color: { rgb: "274E13" } },
                fill: { fgColor: { rgb: "D9EAD3" } },
                alignment: { horizontal: 'center', vertical: 'center' },
            };
        }
    }

    // Data cells (patient products)
    // Columns A (patient), B (product), C (quantity) = white background
    // Column D (price) = green background #E8F5E9
    const dataCellStyleWhite = {
        fill: { fgColor: { rgb: "FFFFFF" } },
        alignment: { horizontal: 'center', vertical: 'center' },
    };
    const dataCellStyleGreen = {
        fill: { fgColor: { rgb: "E8F5E9" } },
        alignment: { horizontal: 'center', vertical: 'center' },
        numFmt: '0.00',
    };

    for (let r = 5; r < sheetData.length; r++) {
        if (orderTotalRows.includes(r) || r === grandTotalRow || sheetData[r].every(v => v === null)) continue;
        for (let c = 0; c < 4; c++) {
            const cellRef = XLSX.utils.encode_cell({ r, c });
            if (ws[cellRef]) {
                ws[cellRef].s = c < 3 ? dataCellStyleWhite : dataCellStyleGreen;
            }
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
            };
        }
        // Ensure merged empty cells B and C also get background
        for (let c = 1; c <= 2; c++) {
            const ref = XLSX.utils.encode_cell({ r: row, c });
            if (!ws[ref]) ws[ref] = { t: 's', v: '' };
            ws[ref].s = {
                fill: { fgColor: { rgb: "FEF5E7" } },
            };
        }
        const valueRef = XLSX.utils.encode_cell({ r: row, c: 3 });
        if (ws[valueRef]) {
            ws[valueRef].s = {
                fill: { fgColor: { rgb: "FEF5E7" } },
                font: { bold: true, color: { rgb: "0000FF" } },
                alignment: { horizontal: 'center', vertical: 'center' },
                numFmt: '0.00',
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

export const generateOrderWordDocument = async (
    comanda: Comanda,
    doctor: Doctor | undefined,
    pacient: Pacient | undefined,
    produse: Produs[]
) => {
    const doc = new Document({
        sections: [{
            properties: {},
            children: [
                new Paragraph({
                    text: 'Fisa Laborator',
                    heading: HeadingLevel.HEADING_1,
                    alignment: AlignmentType.CENTER,
                    spacing: { after: 400 },
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: 'Nume Doctor: ',
                            bold: true,
                        }),
                        new TextRun({
                            text: doctor?.nume || 'N/A',
                        }),
                    ],
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: 'Nume Pacient: ',
                            bold: true,
                        }),
                        new TextRun({
                            text: pacient?.nume || 'N/A',
                        }),
                    ],
                    spacing: { after: 400 },
                }),
                new Paragraph({
                    text: 'Produse:',
                    heading: HeadingLevel.HEADING_2,
                    spacing: { after: 200 },
                }),
                ...comanda.produse.map((comandaProdus) => {
                    const produs = produse.find(p => p.id === comandaProdus.id_produs);
                    return new Paragraph({
                        text: `• ${produs?.nume || 'N/A'} (${comandaProdus.cantitate}x - ${formatCurrency((produs?.pret || 0) * comandaProdus.cantitate)})`,
                        spacing: { after: 100 },
                    });
                }),
                new Paragraph({
                    text: '',
                    spacing: { after: 200 },
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: 'Total: ',
                            bold: true,
                            size: 28,
                        }),
                        new TextRun({
                            text: formatCurrency(comanda.total),
                            size: 28,
                        }),
                    ],
                    spacing: { before: 200 },
                }),
            ],
        }],
    });

    const blob = await Packer.toBlob(doc);
    const filename = `Fisa_Laborator_${doctor?.nume.replace(/\s/g, '_')}_${pacient?.nume.replace(/\s/g, '_')}_${comanda.id}.docx`;
    saveAs(blob, filename);
};
