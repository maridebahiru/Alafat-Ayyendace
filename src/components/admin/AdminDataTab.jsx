import React, { useState } from 'react';
import { Search, Download, UserCheck, BarChart2, Hash, Phone, Clock, Calendar as CalendarIcon, UserPlus, ChevronDown, X, FileSpreadsheet, FileText, Filter } from 'lucide-react';
import { exportToCSV, exportToExcel, exportToPDF, exportTotalAttendedExcel } from '../../utils/exportUtils';
import { isSuperAdmin } from '../../utils/rbac';

const AdminDataTab = ({ 
    students, 
    attendance, 
    setSelectedStudentForAnalytics, 
    manualCheckIn 
}) => {
    const [searchTerm, setSearchTerm] = useState('');
    const [sortConfig, setSortConfig] = useState({ key: 'name', direction: 'asc' });

    // Filter states
    const [startDate, setStartDate] = useState('');
    const [endDate, setEndDate] = useState('');
    const [selectedDepartment, setSelectedDepartment] = useState('All');

    const superAdmin = isSuperAdmin();

    // Get unique departments list
    const departmentsList = ['All', ...new Set(students.map(s => s.department || 'General').filter(Boolean))];

    const handleSort = (key) => {
        setSortConfig(prev => ({
            key,
            direction: prev.key === key && prev.direction === 'asc' ? 'desc' : 'asc'
        }));
    };

    const actionBtnStyle = {
        padding: '8px 14px', borderRadius: '10px', border: '1px solid var(--glass-border)',
        backgroundColor: 'rgba(211, 162, 0, 0.08)', color: 'var(--text-primary)', cursor: 'pointer',
        display: 'flex', alignItems: 'center', gap: '6px', fontSize: '13px', fontWeight: '700',
        transition: 'var(--transition-smooth)'
    };

    const thStyle = { 
        padding: '16px 14px', fontWeight: '700', fontSize: '11px', textTransform: 'uppercase', 
        letterSpacing: '1px', color: 'var(--accent-gold)', borderBottom: '2px solid var(--accent-red)',
        position: 'sticky', top: 0, backgroundColor: 'var(--bg-tertiary)', zIndex: 10
    };
    const tdStyle = { padding: '14px', borderBottom: '1px solid rgba(101, 8, 27, 0.2)', fontSize: '13px' };

    // Filter & Aggregate data
    const filteredStudentsList = students.filter(s => {
        const matchesSearch = s.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
            s.phone.includes(searchTerm) ||
            (s.employeeId && s.employeeId.toLowerCase().includes(searchTerm.toLowerCase())) ||
            (s.idNo && s.idNo.toLowerCase().includes(searchTerm.toLowerCase()));

        const dept = s.department || 'General';
        const matchesDept = selectedDepartment === 'All' || dept === selectedDepartment;

        return matchesSearch && matchesDept;
    });

    const aggregated = filteredStudentsList.map(s => {
        let records = attendance.filter(a => a.phone === s.phone || a.employeeId === s.employeeId);
        if (startDate) records = records.filter(a => a.date >= startDate);
        if (endDate) records = records.filter(a => a.date <= endDate);

        return {
            ...s,
            totalDays: records.length,
            lastSeen: records.sort((a, b) => (b.scannedAt?.seconds || 0) - (a.scannedAt?.seconds || 0))[0]?.scannedAt
        };
    }).sort((a, b) => {
        if (sortConfig.key === 'totalDays') {
            return sortConfig.direction === 'asc' ? a.totalDays - b.totalDays : b.totalDays - a.totalDays;
        }
        const innerRes = a.name.localeCompare(b.name);
        return sortConfig.direction === 'asc' ? innerRes : -innerRes;
    });

    const handleExportCSV = () => exportToCSV({ students, attendance, startDate, endDate, department: selectedDepartment });
    const handleExportExcel = () => exportToExcel({ students, attendance, startDate, endDate, department: selectedDepartment });
    const handleExportTotalAttendedExcel = () => exportTotalAttendedExcel({ students, attendance, startDate, endDate, department: selectedDepartment });
    const handleExportPDF = () => exportToPDF({ students, attendance, startDate, endDate, department: selectedDepartment });

    return (
        <div className="animate-fade-in" style={{ display: 'flex', flexDirection: 'column', gap: window.innerWidth <= 768 ? '16px' : '24px' }}>
            {/* Header Section */}
            <div className="glass-effect" style={{ padding: window.innerWidth <= 768 ? '16px' : '24px', borderRadius: 'var(--radius-lg)', display: 'flex', justifyContent: 'space-between', alignItems: 'center', flexWrap: 'wrap', gap: '20px' }}>
                <div>
                    <h3 style={{ color: 'var(--accent-gold)', fontSize: window.innerWidth <= 768 ? '18px' : '20px', margin: 0 }}>
                        {superAdmin ? "Student Directory & Export Center" : "Student Directory"}
                    </h3>
                    <p style={{ color: 'var(--text-secondary)', fontSize: '13px', marginTop: '4px' }}>
                        Browse {students.length} registered students{superAdmin ? " and export detailed attendance reports." : "."}
                    </p>
                </div>

                {/* Multi-Format Export Buttons (Super Admin Only) */}
                {superAdmin && (
                    <div style={{ display: 'flex', alignItems: 'center', gap: '10px', flexWrap: 'wrap' }}>
                        <button onClick={handleExportCSV} style={actionBtnStyle} title="Export CSV Pivot">
                            <Download size={15} /> CSV
                        </button>
                        <button onClick={handleExportExcel} style={{ ...actionBtnStyle, backgroundColor: 'rgba(0, 255, 128, 0.1)', color: '#00ff80', border: '1px solid rgba(0, 255, 128, 0.3)' }} title="Export Pivot Excel Matrix">
                            <FileSpreadsheet size={15} /> Pivot Excel
                        </button>
                        <button onClick={handleExportTotalAttendedExcel} style={{ ...actionBtnStyle, backgroundColor: 'rgba(211, 162, 0, 0.15)', color: 'var(--accent-gold)', border: '1px solid rgba(211, 162, 0, 0.4)' }} title="Export Total Attended Summary & Scans Report (.xlsx)">
                            <FileSpreadsheet size={15} /> Total Attended (Excel)
                        </button>
                        <button onClick={handleExportPDF} style={{ ...actionBtnStyle, backgroundColor: 'rgba(255, 77, 77, 0.1)', color: '#ff4d4d', border: '1px solid rgba(255, 77, 77, 0.3)' }} title="Export PDF Pivot">
                            <FileText size={15} /> PDF
                        </button>
                    </div>
                )}
            </div>

            {/* Filter Bar */}
            <div className="glass-effect" style={{ padding: '16px 20px', borderRadius: 'var(--radius-lg)', display: 'flex', alignItems: 'center', flexWrap: 'wrap', gap: '16px', fontSize: '13px' }}>
                <div style={{ display: 'flex', alignItems: 'center', gap: '6px', color: 'var(--accent-gold)', fontWeight: '700' }}>
                    <Filter size={16} /> Filters:
                </div>

                {/* Department Dropdown */}
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <span style={{ color: 'var(--text-secondary)' }}>Department:</span>
                    <select
                        value={selectedDepartment}
                        onChange={(e) => setSelectedDepartment(e.target.value)}
                        style={{
                            padding: '8px 12px', borderRadius: '8px', border: '1px solid var(--glass-border)',
                            backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', outline: 'none'
                        }}
                    >
                        {departmentsList.map(dept => (
                            <option key={dept} value={dept}>{dept}</option>
                        ))}
                    </select>
                </div>

                {/* Date Range Start */}
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <span style={{ color: 'var(--text-secondary)' }}>From:</span>
                    <input
                        type="date"
                        value={startDate}
                        onChange={(e) => setStartDate(e.target.value)}
                        style={{
                            padding: '6px 10px', borderRadius: '8px', border: '1px solid var(--glass-border)',
                            backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', outline: 'none'
                        }}
                    />
                </div>

                {/* Date Range End */}
                <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
                    <span style={{ color: 'var(--text-secondary)' }}>To:</span>
                    <input
                        type="date"
                        value={endDate}
                        onChange={(e) => setEndDate(e.target.value)}
                        style={{
                            padding: '6px 10px', borderRadius: '8px', border: '1px solid var(--glass-border)',
                            backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', outline: 'none'
                        }}
                    />
                </div>

                {(startDate || endDate || selectedDepartment !== 'All') && (
                    <button
                        onClick={() => { setStartDate(''); setEndDate(''); setSelectedDepartment('All'); }}
                        style={{ background: 'none', border: 'none', color: '#ff4d4d', cursor: 'pointer', fontWeight: '700', fontSize: '12px' }}
                    >
                        Clear Filters
                    </button>
                )}
            </div>

            {/* Main Table Content */}
            <div className="glass-effect animate-slide-up" style={{ padding: '0', borderRadius: 'var(--radius-lg)', overflow: 'hidden' }}>
                <div style={{ padding: window.innerWidth <= 768 ? '16px' : '24px', backgroundColor: 'rgba(255, 255, 255, 0.02)', borderBottom: '1px solid var(--glass-border)' }}>
                    <div style={{ position: 'relative', maxWidth: '100%' }}>
                        <Search size={20} style={{ position: 'absolute', left: '14px', top: '50%', transform: 'translateY(-50%)', color: 'var(--accent-gold)', opacity: 0.5 }} />
                        <input
                            type="text"
                            placeholder="Find student by name, phone, or ID..."
                            value={searchTerm}
                            onChange={(e) => setSearchTerm(e.target.value)}
                            style={{
                                width: '100%', padding: '12px 12px 12px 44px', borderRadius: '12px', border: '1px solid var(--glass-border)',
                                backgroundColor: 'var(--bg-primary)', color: 'var(--text-primary)', outline: 'none',
                                fontSize: '14px', transition: 'var(--transition-smooth)'
                            }}
                        />
                    </div>
                </div>

                <div className="mobile-scroll-x" style={{ overflowX: 'auto', maxHeight: '700px' }}>
                    <table style={{ width: '100%', borderCollapse: 'collapse', textAlign: 'left', minWidth: '850px' }}>
                        <thead>
                            <tr style={{ cursor: 'pointer' }}>
                                <th style={thStyle} onClick={() => handleSort('name')}>
                                    Student Profile
                                </th>
                                <th style={thStyle}>ID</th>
                                <th style={thStyle}>Department & Position</th>
                                <th style={{ ...thStyle, textAlign: 'center' }} onClick={() => handleSort('totalDays')}>
                                    Total Attended
                                </th>
                                <th style={{ ...thStyle, textAlign: 'center' }}>Management</th>
                            </tr>
                        </thead>
                        <tbody>
                            {aggregated.map((s, i) => (
                                <tr key={s.employeeId || s.phone || i} style={{ borderBottom: '1px solid rgba(101, 8, 27, 0.2)' }} className="table-row-hover">
                                    <td style={tdStyle}>
                                        <div style={{ display: 'flex', alignItems: 'center', gap: '14px' }}>
                                            {s.profilePhotoUrl ? (
                                                <img
                                                    src={s.profilePhotoUrl}
                                                    alt={s.name}
                                                    style={{ width: '42px', height: '42px', borderRadius: '50%', objectFit: 'cover', border: '2px solid var(--accent-gold)' }}
                                                    onError={(e) => { e.target.style.display = 'none'; }}
                                                />
                                            ) : (
                                                <div style={{
                                                    width: '42px', height: '42px', borderRadius: '50%', backgroundColor: 'var(--accent-red)',
                                                    display: 'flex', alignItems: 'center', justifyContent: 'center', color: 'var(--accent-gold)',
                                                    fontSize: '15px', fontWeight: '800'
                                                }}>
                                                    {s.name ? s.name.charAt(0) : '?'}
                                                </div>
                                            )}
                                            <div>
                                                <div style={{ fontWeight: '700', color: 'var(--text-primary)' }}>{s.name}</div>
                                                <div style={{ fontSize: '11px', color: 'var(--text-secondary)', display: 'flex', alignItems: 'center', gap: '4px' }}>
                                                    <Phone size={10} /> {s.phone}
                                                </div>
                                            </div>
                                        </div>
                                    </td>
                                    <td style={tdStyle}>
                                        <span style={{ color: 'var(--accent-gold)', fontWeight: '700', backgroundColor: 'rgba(211, 162, 0, 0.1)', padding: '3px 8px', borderRadius: '6px', fontSize: '12px' }}>
                                            {s.employeeId || s.idNo || 'N/A'}
                                        </span>
                                    </td>
                                    <td style={tdStyle}>
                                        <div style={{ fontWeight: '600', color: 'var(--text-primary)' }}>{s.department || 'General'}</div>
                                        {s.position && <div style={{ fontSize: '11px', color: 'var(--text-secondary)' }}>{s.position}</div>}
                                    </td>
                                    <td style={{ ...tdStyle, textAlign: 'center', fontWeight: '800', color: 'var(--accent-gold)' }}>
                                        {s.totalDays} Days
                                    </td>
                                    <td style={{ ...tdStyle, textAlign: 'center' }}>
                                        <div style={{ display: 'flex', justifyContent: 'center', gap: '8px' }}>
                                            <button
                                                onClick={() => manualCheckIn(s)}
                                                style={{
                                                    padding: '6px 12px', borderRadius: '6px', border: 'none',
                                                    backgroundColor: 'var(--accent-gold)', color: '#1a0a0f',
                                                    fontWeight: '700', fontSize: '12px', cursor: 'pointer'
                                                }}
                                            >
                                                Check In
                                            </button>
                                            <button
                                                onClick={() => setSelectedStudentForAnalytics(s)}
                                                style={{
                                                    padding: '6px 10px', borderRadius: '6px', border: '1px solid var(--glass-border)',
                                                    backgroundColor: 'transparent', color: 'var(--text-primary)',
                                                    fontWeight: '600', fontSize: '12px', cursor: 'pointer'
                                                }}
                                            >
                                                Analytics
                                            </button>
                                        </div>
                                    </td>
                                </tr>
                            ))}
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    );
};

export default AdminDataTab;
