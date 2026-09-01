import ExcelJS from 'exceljs';
import { jsPDF } from 'jspdf';
import html2canvas from 'html2canvas';
import { isSuperAdmin, logSuperAdminAudit } from './rbac';

/**
 * Safely extract millisecond timestamp from Firestore Timestamp, Date, string, or number.
 */
const getTimestampMs = (scannedAt) => {
  if (!scannedAt) return null;
  if (typeof scannedAt.toDate === 'function') return scannedAt.toDate().getTime();
  if (scannedAt instanceof Date) return scannedAt.getTime();
  if (typeof scannedAt === 'object' && typeof scannedAt.seconds === 'number') return scannedAt.seconds * 1000;
  if (typeof scannedAt === 'number') return scannedAt;
  if (typeof scannedAt === 'string') {
    const parsed = Date.parse(scannedAt);
    if (!isNaN(parsed)) return parsed;
  }
  return null;
};

/**
 * Format timestamp into HH:mm:ss string (e.g. 16:42:09).
 */
const formatTimeStr = (scannedAt) => {
  const ms = getTimestampMs(scannedAt);
  if (ms === null) {
    if (typeof scannedAt === 'string' && scannedAt.trim() !== '') return scannedAt;
    return 'Scanned';
  }
  const d = new Date(ms);
  const hours = String(d.getHours()).padStart(2, '0');
  const mins = String(d.getMinutes()).padStart(2, '0');
  const secs = String(d.getSeconds()).padStart(2, '0');
  return `${hours}:${mins}:${secs}`;
};

/**
 * Format date string to YYYY-MM-DD.
 */
const formatDateStr = (dateInput) => {
  if (!dateInput) return '';
  if (typeof dateInput === 'string') {
    if (/^\d{4}-\d{2}-\d{2}$/.test(dateInput)) return dateInput;
    const d = new Date(dateInput);
    if (!isNaN(d.getTime())) return d.toISOString().split('T')[0];
    return dateInput;
  }
  const ms = getTimestampMs(dateInput);
  if (ms !== null) return new Date(ms).toISOString().split('T')[0];
  return '';
};

/**
 * Helper to generate all YYYY-MM-DD date strings between start and end inclusive.
 */
const generateDateRange = (startStr, endStr) => {
  const start = new Date(startStr);
  const end = new Date(endStr);
  if (isNaN(start.getTime()) || isNaN(end.getTime()) || start > end) {
    return [startStr];
  }
  const dates = [];
  const curr = new Date(start);
  while (curr <= end) {
    dates.push(curr.toISOString().split('T')[0]);
    curr.setDate(curr.getDate() + 1);
  }
  return dates;
};

/**
 * Single Data Prep Function for both Excel & PDF exports.
 * Builds structured pivot data with lookup of (student + date) -> earliest scan time.
 */
export const preparePivotReportData = ({ students = [], attendance = [], startDate = '', endDate = '', department = 'All' }) => {
  // 1. Filter students by department
  let filteredStudents = [...students];
  if (department && department !== 'All') {
    filteredStudents = filteredStudents.filter(s => (s.department || 'General') === department);
  }

  // Sort students alphabetically by name using localeCompare with Ethiopic support
  filteredStudents.sort((a, b) => (a.name || '').localeCompare(b.name || '', 'am'));

  // 2. Determine target date range
  let dates = [];
  if (startDate && endDate) {
    dates = generateDateRange(startDate, endDate);
  } else {
    // Extract unique dates from attendance records
    const attendanceDates = new Set();
    attendance.forEach(a => {
      const dStr = formatDateStr(a.date) || formatDateStr(a.scannedAt);
      if (dStr) {
        if (startDate && dStr < startDate) return;
        if (endDate && dStr > endDate) return;
        attendanceDates.add(dStr);
      }
    });

    dates = Array.from(attendanceDates).sort();
    if (dates.length === 0) {
      // Fallback to today if no date records found
      dates = [new Date().toISOString().split('T')[0]];
    }
  }

  // 3. Build lookup map: (studentIdentifier + date) -> { earliestMs, timeStr }
  const scanLookup = new Map();

  attendance.forEach(a => {
    const aDate = formatDateStr(a.date) || formatDateStr(a.scannedAt);
    if (!aDate || !dates.includes(aDate)) return;

    const ms = getTimestampMs(a.scannedAt) || 0;
    const timeStr = formatTimeStr(a.scannedAt);

    // Track by phone, employeeId, and idNo to ensure matching regardless of key used
    const keys = [];
    if (a.phone) keys.push(`${a.phone}_${aDate}`);
    if (a.employeeId) keys.push(`${a.employeeId}_${aDate}`);

    keys.forEach(key => {
      const existing = scanLookup.get(key);
      if (!existing || ms < existing.ms) {
        scanLookup.set(key, { ms, timeStr });
      }
    });
  });

  // 4. Construct pivot headers and student rows
  const title = "Attendance Report — Name / Date / Scan Time";
  const startDesc = startDate || dates[0] || 'Start';
  const endDesc = endDate || dates[dates.length - 1] || 'End';
  const subtitle = `Date Range: ${startDesc} to ${endDesc} | Total Students: ${filteredStudents.length}${department && department !== 'All' ? ` | Department: ${department}` : ''}`;
  const headers = ["#", "Name", "Employee ID", "Phone", ...dates];

  const studentRows = filteredStudents.map((s, idx) => {
    const phone = s.phone || '';
    const empId = s.employeeId || s.idNo || '';

    const dateValues = dates.map(date => {
      const byPhone = phone ? scanLookup.get(`${phone}_${date}`) : null;
      const byEmpId = empId ? scanLookup.get(`${empId}_${date}`) : null;

      const record = byPhone || byEmpId;
      if (record && record.timeStr) {
        return record.timeStr;
      }
      return "-";
    });

    return {
      index: idx + 1,
      name: s.name || 'N/A',
      employeeId: empId || 'N/A',
      phone: phone || 'N/A',
      department: s.department || 'General',
      dateValues
    };
  });

  return { title, subtitle, dates, headers, studentRows };
};

/**
 * Export Pivot Report as formatted Excel (.xlsx) using ExcelJS.
 * Uses Segoe UI / Ethiopic Unicode font family so Amharic names are rendered without distortion.
 */
export const exportToExcel = async ({ students, attendance, startDate, endDate, department }) => {
  const pivotData = preparePivotReportData({ students, attendance, startDate, endDate, department });

  if (pivotData.studentRows.length === 0) {
    alert("No student records found matching the selected filters.");
    return;
  }

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'EJAT Attendance System';
  workbook.created = new Date();

  const sheet = workbook.addWorksheet('Attendance Pivot Report', {
    views: [{ showGridLines: true }]
  });

  const fontFamily = 'Segoe UI';

  // Title Row (Row 1)
  const titleRow = sheet.addRow([pivotData.title]);
  titleRow.height = 30;
  sheet.mergeCells(1, 1, 1, pivotData.headers.length);
  const titleCell = titleRow.getCell(1);
  titleCell.font = { name: fontFamily, size: 16, bold: true, color: { argb: 'FF65081B' } };
  titleCell.alignment = { vertical: 'middle', horizontal: 'left' };

  // Subtitle Row (Row 2)
  const subtitleRow = sheet.addRow([pivotData.subtitle]);
  subtitleRow.height = 20;
  sheet.mergeCells(2, 1, 2, pivotData.headers.length);
  const subtitleCell = subtitleRow.getCell(1);
  subtitleCell.font = { name: fontFamily, size: 10, italic: true, color: { argb: 'FF4B5563' } };
  subtitleCell.alignment = { vertical: 'middle', horizontal: 'left' };

  // Spacing Row (Row 3)
  sheet.addRow([]);

  // Header Row (Row 4)
  const headerRow = sheet.addRow(pivotData.headers);
  headerRow.height = 26;

  headerRow.eachCell((cell, colNumber) => {
    cell.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF65081B' } // Brand color #65081B background
    };
    cell.font = {
      name: fontFamily,
      size: 11,
      bold: true,
      color: { argb: 'FFFFFFFF' } // White text
    };
    cell.alignment = {
      vertical: 'middle',
      horizontal: colNumber === 2 ? 'left' : 'center'
    };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FF65081B' } },
      left: { style: 'thin', color: { argb: 'FF8B1B32' } },
      bottom: { style: 'medium', color: { argb: 'FF65081B' } },
      right: { style: 'thin', color: { argb: 'FF8B1B32' } }
    };
  });

  // Thin border helper for data cells
  const cellBorder = {
    top: { style: 'thin', color: { argb: 'FFE5E7EB' } },
    left: { style: 'thin', color: { argb: 'FFE5E7EB' } },
    bottom: { style: 'thin', color: { argb: 'FFE5E7EB' } },
    right: { style: 'thin', color: { argb: 'FFE5E7EB' } }
  };

  // Student Data Rows (Row 5+)
  pivotData.studentRows.forEach(s => {
    const rowValues = [s.index, s.name, s.employeeId, s.phone, ...s.dateValues];
    const dataRow = sheet.addRow(rowValues);
    dataRow.height = 22;

    dataRow.eachCell((cell, colNumber) => {
      cell.border = cellBorder;
      cell.font = { name: fontFamily, size: 10 };

      if (colNumber === 1) {
        // # column
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.font = { name: fontFamily, size: 10, color: { argb: 'FF6B7280' } };
      } else if (colNumber === 2) {
        // Name column (Amharic UTF-8 string)
        cell.alignment = { vertical: 'middle', horizontal: 'left' };
        cell.font = { name: fontFamily, size: 10, bold: true, color: { argb: 'FF111827' } };
      } else if (colNumber === 3 || colNumber === 4) {
        // Employee ID & Phone
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.font = { name: fontFamily, size: 10, color: { argb: 'FF374151' } };
      } else {
        // Date columns (Col 5+)
        const val = String(cell.value || '').trim();
        if (val !== '-') {
          // Present: Light green background, dark green bold text
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFDCFCE7' }
          };
          cell.font = {
            name: fontFamily,
            size: 10,
            bold: true,
            color: { argb: 'FF166534' }
          };
          cell.alignment = { vertical: 'middle', horizontal: 'center' };
        } else {
          // Absent: Light red/pink background, dark red text
          cell.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFFEE2E2' }
          };
          cell.font = {
            name: fontFamily,
            size: 10,
            color: { argb: 'FF991B1B' }
          };
          cell.alignment = { vertical: 'middle', horizontal: 'center' };
        }
      }
    });
  });

  // Set column widths cleanly
  sheet.getColumn(1).width = 6;   // #
  sheet.getColumn(2).width = 30;  // Name (wider for Amharic text)
  sheet.getColumn(3).width = 18;  // Employee ID
  sheet.getColumn(4).width = 16;  // Phone

  for (let i = 5; i <= pivotData.headers.length; i++) {
    sheet.getColumn(i).width = 15; // Date columns
  }

  // Generate buffer and trigger browser download
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const fileDate = new Date().toISOString().split('T')[0];
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = `Attendance_Report_Pivot_${fileDate}.xlsx`;
  link.click();
  URL.revokeObjectURL(link.href);
};

/**
 * Export Pivot Report as PDF using HTML-to-Canvas rendering with jsPDF.
 * Renders native browser DOM fonts (Segoe UI / Ethiopic Unicode) so Amharic names
 * are rendered 100% crystal clear without distortion, using brand color #65081B.
 */
export const exportToPDF = async ({ students, attendance, startDate, endDate, department }) => {
  const pivotData = preparePivotReportData({ students, attendance, startDate, endDate, department });

  if (pivotData.studentRows.length === 0) {
    alert("No student records found matching the selected filters.");
    return;
  }

  // Create off-screen HTML element for exact rendering with Ethiopic font support & brand color #65081B
  const container = document.createElement('div');
  container.style.position = 'absolute';
  container.style.left = '-9999px';
  container.style.top = '-9999px';
  container.style.width = '1200px';
  container.style.backgroundColor = '#ffffff';
  container.style.padding = '28px';
  container.style.boxSizing = 'border-box';
  container.style.fontFamily = "'Segoe UI', 'Nyala', 'Abyssinica SIL', 'Noto Sans Ethiopic', Arial, sans-serif";

  // Build header row HTML with brand color #65081B
  const tableHeadersHtml = pivotData.headers.map((h, i) =>
    `<th style="background-color: #65081B; color: #ffffff; padding: 10px 8px; font-size: 13px; font-weight: bold; text-align: ${i === 1 ? 'left' : 'center'}; border: 1px solid #65081B; white-space: nowrap;">${h}</th>`
  ).join('');

  // Build student rows HTML
  const tableRowsHtml = pivotData.studentRows.map(s => {
    const dateColsHtml = s.dateValues.map(val => {
      const isPresent = val !== '-';
      const bgColor = isPresent ? '#DCFCE7' : '#FEE2E2';
      const textColor = isPresent ? '#166534' : '#991B1B';
      const fontWeight = isPresent ? 'bold' : 'normal';
      return `<td style="background-color: ${bgColor}; color: ${textColor}; padding: 8px 6px; font-size: 12px; font-weight: ${fontWeight}; text-align: center; border: 1px solid #e5e7eb; white-space: nowrap;">${val}</td>`;
    }).join('');

    return `
      <tr>
        <td style="padding: 8px 6px; font-size: 12px; text-align: center; border: 1px solid #e5e7eb; color: #6b7280;">${s.index}</td>
        <td style="padding: 8px 10px; font-size: 12px; text-align: left; font-weight: bold; border: 1px solid #e5e7eb; color: #111827; white-space: nowrap;">${s.name}</td>
        <td style="padding: 8px 6px; font-size: 12px; text-align: center; border: 1px solid #e5e7eb; color: #374151; white-space: nowrap;">${s.employeeId}</td>
        <td style="padding: 8px 6px; font-size: 12px; text-align: center; border: 1px solid #e5e7eb; color: #374151; white-space: nowrap;">${s.phone}</td>
        ${dateColsHtml}
      </tr>
    `;
  }).join('');

  container.innerHTML = `
    <div style="margin-bottom: 20px;">
      <h1 style="color: #65081B; font-size: 24px; font-weight: bold; margin: 0 0 6px 0; font-family: 'Segoe UI', 'Nyala', 'Noto Sans Ethiopic', sans-serif;">${pivotData.title}</h1>
      <p style="color: #4b5563; font-size: 13px; margin: 0; font-style: italic;">${pivotData.subtitle}</p>
    </div>
    <table style="width: 100%; border-collapse: collapse; margin-top: 12px; background-color: #ffffff;">
      <thead>
        <tr>${tableHeadersHtml}</tr>
      </thead>
      <tbody>
        ${tableRowsHtml}
      </tbody>
    </table>
  `;

  document.body.appendChild(container);

  try {
    const canvas = await html2canvas(container, {
      scale: 2,
      useCORS: true,
      logging: false,
      backgroundColor: '#ffffff'
    });

    document.body.removeChild(container);

    const imgData = canvas.toDataURL('image/png');
    const pdf = new jsPDF({
      orientation: 'landscape',
      unit: 'mm',
      format: 'a4'
    });

    const pdfWidth = pdf.internal.pageSize.getWidth();
    const pdfHeight = pdf.internal.pageSize.getHeight();

    const imgWidth = pdfWidth - 20; // 10mm margins
    const imgHeight = (canvas.height * imgWidth) / canvas.width;

    if (imgHeight <= pdfHeight - 20) {
      pdf.addImage(imgData, 'PNG', 10, 10, imgWidth, imgHeight);
    } else {
      // Multi-page slicing if table has many rows
      let heightLeft = imgHeight;
      let position = 10;

      pdf.addImage(imgData, 'PNG', 10, position, imgWidth, imgHeight);
      heightLeft -= (pdfHeight - 20);

      while (heightLeft > 0) {
        position = heightLeft - imgHeight + 10;
        pdf.addPage();
        pdf.addImage(imgData, 'PNG', 10, position, imgWidth, imgHeight);
        heightLeft -= (pdfHeight - 20);
      }
    }

    const fileDate = new Date().toISOString().split('T')[0];
    pdf.save(`Attendance_Report_Pivot_${fileDate}.pdf`);
  } catch (err) {
    console.error('Error generating PDF:', err);
    if (document.body.contains(container)) {
      document.body.removeChild(container);
    }
    alert('Failed to generate PDF report. Please try again.');
  }
};

/**
 * Export Pivot Report as CSV.
 */
export const exportToCSV = ({ students, attendance, startDate, endDate, department }) => {
  const pivotData = preparePivotReportData({ students, attendance, startDate, endDate, department });

  if (pivotData.studentRows.length === 0) {
    alert("No student records found matching the selected filters.");
    return;
  }

  const csvRows = [
    [pivotData.title],
    [pivotData.subtitle],
    [],
    pivotData.headers.map(h => `"${String(h).replace(/"/g, '""')}"`).join(',')
  ];

  pivotData.studentRows.forEach(r => {
    const row = [r.index, r.name, r.employeeId, r.phone, ...r.dateValues];
    csvRows.push(row.map(val => `"${String(val).replace(/"/g, '""')}"`).join(','));
  });

  const csvContent = csvRows.join('\n');
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const fileDate = new Date().toISOString().split('T')[0];
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = `Attendance_Report_Pivot_${fileDate}.csv`;
  link.click();
  URL.revokeObjectURL(link.href);
};

/**
 * Data Prep Function for Total Attended Summary & Detailed Logs Report (Super Admin).
 */
export const prepareTotalAttendedReportData = ({ students = [], attendance = [], startDate = '', endDate = '', department = 'All' }) => {
  // 1. Filter students by department
  let filteredStudents = [...students];
  if (department && department !== 'All') {
    filteredStudents = filteredStudents.filter(s => (s.department || 'General') === department);
  }

  // Sort students alphabetically by name using localeCompare with Ethiopic support
  filteredStudents.sort((a, b) => (a.name || '').localeCompare(b.name || '', 'am'));

  // 2. Filter attendance records by date range
  let filteredAttendance = [...attendance];
  if (startDate) {
    filteredAttendance = filteredAttendance.filter(a => {
      const d = formatDateStr(a.date) || formatDateStr(a.scannedAt);
      return d && d >= startDate;
    });
  }
  if (endDate) {
    filteredAttendance = filteredAttendance.filter(a => {
      const d = formatDateStr(a.date) || formatDateStr(a.scannedAt);
      return d && d <= endDate;
    });
  }

  // Determine total unique event dates in period
  const uniqueDatesInPeriod = new Set(
    filteredAttendance
      .map(a => formatDateStr(a.date) || formatDateStr(a.scannedAt))
      .filter(Boolean)
  );
  const totalEventDays = uniqueDatesInPeriod.size || (startDate && endDate ? generateDateRange(startDate, endDate).length : 1);

  // 3. Aggregate total attendance per student
  const summaryRows = filteredStudents.map((s, idx) => {
    const phone = s.phone || '';
    const empId = s.employeeId || s.idNo || '';

    // Match records for this student
    const studentRecords = filteredAttendance.filter(a => 
      (phone && a.phone === phone) || 
      (empId && (a.employeeId === empId || a.idNo === empId))
    );

    // Get unique dates attended by student
    const studentAttendedDates = new Set(
      studentRecords
        .map(a => formatDateStr(a.date) || formatDateStr(a.scannedAt))
        .filter(Boolean)
    );
    const totalDaysAttended = studentAttendedDates.size;

    // Attendance Rate %
    const rateVal = totalEventDays > 0 ? Math.min(100, Math.round((totalDaysAttended / totalEventDays) * 100)) : 0;

    // Sort timestamps to get first and latest scan
    const timestamps = studentRecords
      .map(a => getTimestampMs(a.scannedAt) || getTimestampMs(a.date))
      .filter(Boolean)
      .sort((a, b) => a - b);

    const firstScanStr = timestamps.length > 0 
      ? new Date(timestamps[0]).toLocaleString('en-US', { dateStyle: 'short', timeStyle: 'short' })
      : 'N/A';
    const latestScanStr = timestamps.length > 0
      ? new Date(timestamps[timestamps.length - 1]).toLocaleString('en-US', { dateStyle: 'short', timeStyle: 'short' })
      : 'N/A';

    let status = 'No Attendance';
    if (totalDaysAttended > 0) {
      if (rateVal >= 80) status = 'Excellent (≥80%)';
      else if (rateVal >= 50) status = 'Good (50-79%)';
      else status = 'Low Attendance (<50%)';
    }

    return {
      index: idx + 1,
      name: s.name || 'N/A',
      employeeId: empId || 'N/A',
      phone: phone || 'N/A',
      department: s.department || 'General',
      position: s.position || 'N/A',
      totalDaysAttended,
      totalEventDays,
      attendanceRate: `${rateVal}%`,
      firstScan: firstScanStr,
      latestScan: latestScanStr,
      status
    };
  });

  // 4. Construct Detailed Log Rows
  const detailedLogs = filteredAttendance.map((a, idx) => {
    const student = students.find(s => 
      (a.phone && s.phone === a.phone) || 
      (a.employeeId && (s.employeeId === a.employeeId || s.idNo === a.employeeId))
    );
    return {
      index: idx + 1,
      date: formatDateStr(a.date) || formatDateStr(a.scannedAt) || 'N/A',
      time: formatTimeStr(a.scannedAt),
      name: student?.name || a.studentName || 'N/A',
      employeeId: a.employeeId || a.idNo || student?.employeeId || 'N/A',
      phone: a.phone || student?.phone || 'N/A',
      department: a.department || student?.department || 'General',
      scannedBy: a.scannedBy || 'System',
      location: typeof a.scannedLocation === 'string' ? a.scannedLocation : (a.scannedLocation?.city || 'N/A')
    };
  }).sort((a, b) => b.date.localeCompare(a.date));

  const title = "Total Attended Summary & Attendance Log Report";
  const startDesc = startDate || 'Beginning';
  const endDesc = endDate || 'Latest';
  const subtitle = `Date Range: ${startDesc} to ${endDesc} | Total Registered Students: ${filteredStudents.length} | Total Event Days: ${totalEventDays} | Department: ${department}`;

  return { title, subtitle, summaryRows, detailedLogs, totalEventDays, totalScans: filteredAttendance.length };
};

/**
 * Export Total Attended Report as formatted multi-sheet Excel (.xlsx) using ExcelJS.
 * Strictly restricted to Super Admin only.
 */
export const exportTotalAttendedExcel = async ({ students, attendance, startDate, endDate, department }) => {
  if (!isSuperAdmin()) {
    alert("Access Denied: The Total Attended Report export is strictly restricted to Super Admin.");
    return;
  }

  const reportData = prepareTotalAttendedReportData({ students, attendance, startDate, endDate, department });

  if (reportData.summaryRows.length === 0) {
    alert("No student records found matching the selected filters.");
    return;
  }

  logSuperAdminAudit('export_total_attended_excel', {
    startDate,
    endDate,
    department,
    studentCount: reportData.summaryRows.length,
    scanCount: reportData.totalScans
  });

  const workbook = new ExcelJS.Workbook();
  workbook.creator = 'EJAT Attendance System - SuperAdmin';
  workbook.created = new Date();
  const fontFamily = 'Segoe UI';

  // --- SHEET 1: Total Attended Summary ---
  const summarySheet = workbook.addWorksheet('Total Attended Summary', {
    views: [{ showGridLines: true }]
  });

  // Title Row (Row 1)
  const titleRow = summarySheet.addRow([reportData.title]);
  titleRow.height = 32;
  summarySheet.mergeCells(1, 1, 1, 12);
  const titleCell = titleRow.getCell(1);
  titleCell.font = { name: fontFamily, size: 16, bold: true, color: { argb: 'FF65081B' } };
  titleCell.alignment = { vertical: 'middle', horizontal: 'left' };

  // Subtitle Row (Row 2)
  const subtitleRow = summarySheet.addRow([reportData.subtitle]);
  subtitleRow.height = 22;
  summarySheet.mergeCells(2, 1, 2, 12);
  const subtitleCell = subtitleRow.getCell(1);
  subtitleCell.font = { name: fontFamily, size: 10, italic: true, color: { argb: 'FF4B5563' } };
  subtitleCell.alignment = { vertical: 'middle', horizontal: 'left' };

  // Spacing Row (Row 3)
  summarySheet.addRow([]);

  // Headers (Row 4)
  const summaryHeaders = [
    "#", "Student Name", "Employee ID", "Phone", "Department", "Position",
    "Total Attended (Days)", "Total Event Days", "Attendance Rate", "First Scan", "Latest Scan", "Status"
  ];

  const headerRow = summarySheet.addRow(summaryHeaders);
  headerRow.height = 28;
  headerRow.eachCell((cell, colNumber) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF65081B' } };
    cell.font = { name: fontFamily, size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.alignment = { vertical: 'middle', horizontal: colNumber === 2 ? 'left' : 'center' };
    cell.border = {
      top: { style: 'thin', color: { argb: 'FF65081B' } },
      left: { style: 'thin', color: { argb: 'FF8B1B32' } },
      bottom: { style: 'medium', color: { argb: 'FF65081B' } },
      right: { style: 'thin', color: { argb: 'FF8B1B32' } }
    };
  });

  const cellBorder = {
    top: { style: 'thin', color: { argb: 'FFE5E7EB' } },
    left: { style: 'thin', color: { argb: 'FFE5E7EB' } },
    bottom: { style: 'thin', color: { argb: 'FFE5E7EB' } },
    right: { style: 'thin', color: { argb: 'FFE5E7EB' } }
  };

  reportData.summaryRows.forEach(s => {
    const rowValues = [
      s.index, s.name, s.employeeId, s.phone, s.department, s.position,
      s.totalDaysAttended, s.totalEventDays, s.attendanceRate, s.firstScan, s.latestScan, s.status
    ];
    const dataRow = summarySheet.addRow(rowValues);
    dataRow.height = 24;

    dataRow.eachCell((cell, colNumber) => {
      cell.border = cellBorder;
      cell.font = { name: fontFamily, size: 10 };

      if (colNumber === 1) {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.font = { name: fontFamily, size: 10, color: { argb: 'FF6B7280' } };
      } else if (colNumber === 2) {
        cell.alignment = { vertical: 'middle', horizontal: 'left' };
        cell.font = { name: fontFamily, size: 10, bold: true, color: { argb: 'FF111827' } };
      } else if (colNumber === 7) {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.font = { name: fontFamily, size: 11, bold: true, color: { argb: 'FF65081B' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFDF2F8' } };
      } else if (colNumber === 9) {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.font = { name: fontFamily, size: 10, bold: true };
      } else if (colNumber === 12) {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        if (s.status.startsWith('Excellent')) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCFCE7' } };
          cell.font = { name: fontFamily, size: 10, bold: true, color: { argb: 'FF166534' } };
        } else if (s.status.startsWith('Good')) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEF9C3' } };
          cell.font = { name: fontFamily, size: 10, bold: true, color: { argb: 'FF854D0E' } };
        } else if (s.status.startsWith('Low')) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
          cell.font = { name: fontFamily, size: 10, color: { argb: 'FF991B1B' } };
        } else {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3F4F6' } };
          cell.font = { name: fontFamily, size: 10, color: { argb: 'FF6B7280' } };
        }
      } else {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
      }
    });
  });

  summarySheet.getColumn(1).width = 6;   // #
  summarySheet.getColumn(2).width = 28;  // Name
  summarySheet.getColumn(3).width = 16;  // Employee ID
  summarySheet.getColumn(4).width = 16;  // Phone
  summarySheet.getColumn(5).width = 18;  // Department
  summarySheet.getColumn(6).width = 18;  // Position
  summarySheet.getColumn(7).width = 22;  // Total Attended (Days)
  summarySheet.getColumn(8).width = 18;  // Total Event Days
  summarySheet.getColumn(9).width = 18;  // Attendance Rate
  summarySheet.getColumn(10).width = 20; // First Scan
  summarySheet.getColumn(11).width = 20; // Latest Scan
  summarySheet.getColumn(12).width = 24; // Status


  // --- SHEET 2: Detailed Attended Scans Log ---
  const detailSheet = workbook.addWorksheet('Detailed Attended Logs', {
    views: [{ showGridLines: true }]
  });

  const dTitleRow = detailSheet.addRow([`Detailed Scans Log (${reportData.detailedLogs.length} Total Records)`]);
  dTitleRow.height = 28;
  detailSheet.mergeCells(1, 1, 1, 9);
  const dTitleCell = dTitleRow.getCell(1);
  dTitleCell.font = { name: fontFamily, size: 14, bold: true, color: { argb: 'FF65081B' } };
  dTitleCell.alignment = { vertical: 'middle', horizontal: 'left' };

  detailSheet.addRow([]);

  const detailHeaders = ["#", "Date", "Time", "Student Name", "Employee ID", "Phone", "Department", "Scanned By", "Location"];
  const dHeaderRow = detailSheet.addRow(detailHeaders);
  dHeaderRow.height = 26;

  dHeaderRow.eachCell((cell, colNumber) => {
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF65081B' } };
    cell.font = { name: fontFamily, size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.alignment = { vertical: 'middle', horizontal: colNumber === 4 ? 'left' : 'center' };
    cell.border = cellBorder;
  });

  reportData.detailedLogs.forEach(d => {
    const rowValues = [d.index, d.date, d.time, d.name, d.employeeId, d.phone, d.department, d.scannedBy, d.location];
    const dataRow = detailSheet.addRow(rowValues);
    dataRow.height = 22;

    dataRow.eachCell((cell, colNumber) => {
      cell.border = cellBorder;
      cell.font = { name: fontFamily, size: 10 };
      cell.alignment = { vertical: 'middle', horizontal: colNumber === 4 ? 'left' : 'center' };
    });
  });

  detailSheet.getColumn(1).width = 6;   // #
  detailSheet.getColumn(2).width = 14;  // Date
  detailSheet.getColumn(3).width = 12;  // Time
  detailSheet.getColumn(4).width = 28;  // Name
  detailSheet.getColumn(5).width = 16;  // Employee ID
  detailSheet.getColumn(6).width = 16;  // Phone
  detailSheet.getColumn(7).width = 18;  // Department
  detailSheet.getColumn(8).width = 20;  // Scanned By
  detailSheet.getColumn(9).width = 22;  // Location

  // Save & Trigger Download
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const fileDate = new Date().toISOString().split('T')[0];
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = `Total_Attended_Report_${fileDate}.xlsx`;
  link.click();
  URL.revokeObjectURL(link.href);
};

