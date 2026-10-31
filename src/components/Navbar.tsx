import React, { useState, useRef, useEffect } from 'react';
import { NavigationTab } from '../types';
import { DEPARTMENTS } from '../data/defaultData';
import {
  BarChart3,
  Settings as SettingsIcon,
  Users,
  Building2,
  PieChart,
  Clock,
  Download,
  Upload,
  FileSpreadsheet,
  RotateCcw,
  Save,
  ChevronDown,
  Menu,
  X,
  Layers,
  ShieldCheck,
} from 'lucide-react';

interface NavbarProps {
  activeTab: NavigationTab;
  setActiveTab: (tab: NavigationTab) => void;
  lastSaved: string;
  onExportBackup: () => void;
  onOpenImportBackup: () => void;
  onOpenBulkCsv: () => void;
  onOpenResetModal: () => void;
}

export const Navbar: React.FC<NavbarProps> = ({
  activeTab,
  setActiveTab,
  lastSaved,
  onExportBackup,
  onOpenImportBackup,
  onOpenBulkCsv,
  onOpenResetModal,
}) => {
  const isDeptActive = activeTab.startsWith('dept_');
  const [isDeptDropdownOpen, setIsDeptDropdownOpen] = useState(false);
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);
  const deptDropdownRef = useRef<HTMLDivElement>(null);

  const activeDept = DEPARTMENTS.find((d) => d.id === activeTab);

  // Close dropdown on outside click
  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (
        deptDropdownRef.current &&
        !deptDropdownRef.current.contains(event.target as Node)
      ) {
        setIsDeptDropdownOpen(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const handleSelectTab = (tab: NavigationTab) => {
    setActiveTab(tab);
    setIsDeptDropdownOpen(false);
    setIsMobileMenuOpen(false);
  };

  return (
    <>
      {/* 56px Sticky Top Horizontal Navigation Bar */}
      <header className="sticky top-0 z-40 bg-white border-b border-[#E8E8E6] shadow-xs">
        <div className="max-w-[1400px] mx-auto px-4 sm:px-6 lg:px-10 h-14 flex items-center justify-between gap-4">
          
          {/* Top-Left Branding: Theseus Workshop + Product Description */}
          <div className="flex items-center min-w-0 shrink-0 gap-3">
            <div className="flex items-center gap-2">
              <span className="font-garamond text-base sm:text-lg font-bold text-[#051C2C] tracking-tight whitespace-nowrap">
                Theseus Workshop
              </span>
              <span className="hidden sm:inline-block w-1.5 h-1.5 rounded-full bg-[#2251FF]" />
            </div>

            <div className="hidden md:flex items-center border-l border-[#E8E8E6] pl-3 min-w-0">
              <span
                className="text-xs text-[#888888] font-normal truncate max-w-[240px] lg:max-w-xs xl:max-w-md 2xl:max-w-lg"
                title="Enterprise Payroll & Annual Workforce Capacity Planning Excel Toolkit"
              >
                Enterprise Payroll & Annual Workforce Capacity Planning Excel Toolkit
              </span>
            </div>
          </div>

          {/* Desktop Navigation Links */}
          <nav className="hidden lg:flex items-center space-x-1 font-medium text-xs">
            {/* 1. Dashboard */}
            <button
              onClick={() => handleSelectTab('dashboard')}
              className={`px-3 py-1.5 rounded-lg transition-all flex items-center gap-1.5 ${
                activeTab === 'dashboard'
                  ? 'text-[#2251FF] font-semibold bg-[#2251FF]/8 shadow-2xs border-b-2 border-[#2251FF]'
                  : 'text-[#051C2C] hover:text-[#2251FF] hover:bg-gray-100'
              }`}
            >
              <BarChart3 className="w-3.5 h-3.5 shrink-0" />
              <span>Dashboard</span>
            </button>

            {/* 2. Settings */}
            <button
              onClick={() => handleSelectTab('settings')}
              className={`px-3 py-1.5 rounded-lg transition-all flex items-center gap-1.5 ${
                activeTab === 'settings'
                  ? 'text-[#2251FF] font-semibold bg-[#2251FF]/8 shadow-2xs border-b-2 border-[#2251FF]'
                  : 'text-[#051C2C] hover:text-[#2251FF] hover:bg-gray-100'
              }`}
            >
              <SettingsIcon className="w-3.5 h-3.5 shrink-0" />
              <span>Settings</span>
            </button>

            {/* 3. Employees */}
            <button
              onClick={() => handleSelectTab('employee_master')}
              className={`px-3 py-1.5 rounded-lg transition-all flex items-center gap-1.5 ${
                activeTab === 'employee_master'
                  ? 'text-[#2251FF] font-semibold bg-[#2251FF]/8 shadow-2xs border-b-2 border-[#2251FF]'
                  : 'text-[#051C2C] hover:text-[#2251FF] hover:bg-gray-100'
              }`}
            >
              <Users className="w-3.5 h-3.5 shrink-0" />
              <span>Employees</span>
            </button>

            {/* 4. Depts Dropdown */}
            <div className="relative" ref={deptDropdownRef}>
              <button
                onClick={() => setIsDeptDropdownOpen(!isDeptDropdownOpen)}
                className={`px-3 py-1.5 rounded-lg transition-all flex items-center gap-1.5 ${
                  isDeptActive
                    ? 'text-[#2251FF] font-semibold bg-[#2251FF]/8 shadow-2xs border-b-2 border-[#2251FF]'
                    : 'text-[#051C2C] hover:text-[#2251FF] hover:bg-gray-100'
                }`}
              >
                <Building2 className="w-3.5 h-3.5 shrink-0" />
                <span>Depts</span>
                {isDeptActive && activeDept && (
                  <span className="text-[10px] bg-[#2251FF] text-white px-1.5 py-0.2 rounded font-mono font-bold">
                    {activeDept.deptCode}
                  </span>
                )}
                <ChevronDown className={`w-3 h-3 text-[#888888] transition-transform ${isDeptDropdownOpen ? 'rotate-180' : ''}`} />
              </button>

              {isDeptDropdownOpen && (
                <div className="absolute left-0 mt-1.5 w-60 bg-white rounded-xl shadow-lg border border-[#E8E8E6] p-1.5 z-50 animate-in fade-in zoom-in-95 duration-150">
                  <div className="px-2.5 py-1 text-[10px] font-bold text-[#888888] uppercase tracking-wider flex items-center gap-1">
                    <Layers className="w-3 h-3 text-[#2251FF]" />
                    <span>Department Sheets</span>
                  </div>
                  <div className="space-y-0.5 mt-1">
                    {DEPARTMENTS.map((dept) => {
                      const isSelected = activeTab === dept.id;
                      return (
                        <button
                          key={dept.id}
                          onClick={() => handleSelectTab(dept.id as NavigationTab)}
                          className={`w-full text-left px-2.5 py-2 rounded-lg text-xs flex items-center justify-between transition-colors ${
                            isSelected
                              ? 'bg-[#051C2C] text-white font-semibold'
                              : 'text-[#051C2C] hover:bg-gray-100 hover:text-[#2251FF]'
                          }`}
                        >
                          <span className="truncate">{dept.shortName}: {dept.name}</span>
                          <span
                            className={`text-[9px] font-mono font-bold px-1.5 py-0.5 rounded ${
                              isSelected
                                ? 'bg-white/20 text-white'
                                : 'bg-gray-200/80 text-[#051C2C]'
                            }`}
                          >
                            {dept.deptCode}
                          </span>
                        </button>
                      );
                    })}
                  </div>
                </div>
              )}
            </div>

            {/* 5. Payroll Summary */}
            <button
              onClick={() => handleSelectTab('payroll_summary')}
              className={`px-3 py-1.5 rounded-lg transition-all flex items-center gap-1.5 ${
                activeTab === 'payroll_summary'
                  ? 'text-[#2251FF] font-semibold bg-[#2251FF]/8 shadow-2xs border-b-2 border-[#2251FF]'
                  : 'text-[#051C2C] hover:text-[#2251FF] hover:bg-gray-100'
              }`}
            >
              <PieChart className="w-3.5 h-3.5 shrink-0" />
              <span>Payroll Summary</span>
            </button>

            {/* 6. Hours Tracker */}
            <button
              onClick={() => handleSelectTab('annual_hours_tracker')}
              className={`px-3 py-1.5 rounded-lg transition-all flex items-center gap-1.5 ${
                activeTab === 'annual_hours_tracker'
                  ? 'text-[#2251FF] font-semibold bg-[#2251FF]/8 shadow-2xs border-b-2 border-[#2251FF]'
                  : 'text-[#051C2C] hover:text-[#2251FF] hover:bg-gray-100'
              }`}
            >
              <Clock className="w-3.5 h-3.5 shrink-0" />
              <span>Hours Tracker</span>
            </button>
          </nav>

          {/* Right Area: Last Saved & Data Action Buttons */}
          <div className="flex items-center gap-2">
            {/* Last saved status badge */}
            <div className="hidden sm:flex items-center gap-1 text-[11px] text-[#888888] bg-[#F5F5F2] px-2.5 py-1 rounded-lg border border-[#E8E8E6] whitespace-nowrap">
              <Save className="w-3.5 h-3.5 text-[#00C853] shrink-0" />
              <span className="font-medium text-[#051C2C]">Saved:</span>
              <span className="truncate max-w-[110px]">{lastSaved || 'Just now'}</span>
            </div>

            {/* Desktop Action Buttons */}
            <div className="hidden md:flex items-center gap-1.5 border-l border-[#E8E8E6] pl-2">
              <button
                onClick={onExportBackup}
                className="px-2.5 py-1.5 rounded-lg bg-[#F5F5F2] hover:bg-gray-200 text-[#051C2C] text-xs font-medium flex items-center gap-1 transition-colors border border-[#E8E8E6]"
                title="Export complete backup JSON file"
              >
                <Download className="w-3.5 h-3.5 text-[#2251FF]" />
                <span className="hidden xl:inline">Export</span>
              </button>

              <button
                onClick={onOpenImportBackup}
                className="px-2.5 py-1.5 rounded-lg bg-[#F5F5F2] hover:bg-gray-200 text-[#051C2C] text-xs font-medium flex items-center gap-1 transition-colors border border-[#E8E8E6]"
                title="Import and validate backup file"
              >
                <Upload className="w-3.5 h-3.5 text-[#2251FF]" />
                <span className="hidden xl:inline">Import</span>
              </button>

              <button
                onClick={onOpenBulkCsv}
                className="px-2.5 py-1.5 rounded-lg bg-[#F5F5F2] hover:bg-gray-200 text-[#051C2C] text-xs font-medium flex items-center gap-1 transition-colors border border-[#E8E8E6]"
                title="Bulk CSV Import into database tables"
              >
                <FileSpreadsheet className="w-3.5 h-3.5 text-[#00C853]" />
                <span className="hidden xl:inline">Bulk CSV</span>
              </button>

              <button
                onClick={onOpenResetModal}
                className="px-2.5 py-1.5 rounded-lg bg-red-50 hover:bg-red-100 text-[#D32F2F] text-xs font-medium flex items-center gap-1 transition-colors border border-red-200"
                title="Reset all tables to initial workbook state"
              >
                <RotateCcw className="w-3.5 h-3.5 text-[#D32F2F]" />
                <span className="hidden xl:inline">Reset</span>
              </button>
            </div>

            {/* Mobile / Tablet Menu Button */}
            <button
              onClick={() => setIsMobileMenuOpen(!isMobileMenuOpen)}
              className="lg:hidden p-2 rounded-lg bg-[#F5F5F2] hover:bg-[#E8E8E6] text-[#051C2C] transition-colors"
              aria-label="Toggle navigation menu"
            >
              {isMobileMenuOpen ? <X className="w-5 h-5" /> : <Menu className="w-5 h-5" />}
            </button>
          </div>
        </div>
      </header>

      {/* Mobile Drawer Overlay */}
      {isMobileMenuOpen && (
        <div className="lg:hidden fixed inset-0 z-50 bg-[#051C2C]/50 backdrop-blur-xs flex justify-end">
          <div className="w-80 max-w-[85vw] bg-white h-full shadow-2xl flex flex-col justify-between p-5 animate-in slide-in-from-right duration-200">
            <div className="space-y-4 overflow-y-auto pr-1">
              {/* Header */}
              <div className="flex items-center justify-between pb-3 border-b border-[#E8E8E6]">
                <div>
                  <div className="font-garamond text-base font-bold text-[#051C2C]">
                    Theseus Workshop
                  </div>
                  <div className="text-[10px] text-[#888888] leading-tight mt-0.5">
                    Enterprise Payroll & Annual Workforce Capacity Planning Excel Toolkit
                  </div>
                </div>
                <button
                  onClick={() => setIsMobileMenuOpen(false)}
                  className="p-1.5 rounded-lg hover:bg-gray-100 text-[#888888]"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>

              {/* Saved Status */}
              <div className="flex items-center gap-1.5 text-xs text-[#888888] bg-[#F5F5F2] px-3 py-2 rounded-lg border border-[#E8E8E6]">
                <Save className="w-3.5 h-3.5 text-[#00C853] shrink-0" />
                <span>Last saved: {lastSaved || 'Just now'}</span>
              </div>

              {/* Module Links */}
              <div className="space-y-1">
                <div className="px-2 py-1 text-[10px] font-bold text-[#888888] uppercase tracking-wider">
                  Modules
                </div>

                <button
                  onClick={() => handleSelectTab('dashboard')}
                  className={`w-full flex items-center gap-2.5 px-3 py-2 rounded-xl text-xs font-medium ${
                    activeTab === 'dashboard'
                      ? 'bg-[#051C2C] text-white font-semibold'
                      : 'text-[#051C2C] hover:bg-gray-100'
                  }`}
                >
                  <BarChart3 className="w-4 h-4" />
                  <span>Dashboard</span>
                </button>

                <button
                  onClick={() => handleSelectTab('settings')}
                  className={`w-full flex items-center gap-2.5 px-3 py-2 rounded-xl text-xs font-medium ${
                    activeTab === 'settings'
                      ? 'bg-[#051C2C] text-white font-semibold'
                      : 'text-[#051C2C] hover:bg-gray-100'
                  }`}
                >
                  <SettingsIcon className="w-4 h-4" />
                  <span>Settings</span>
                </button>

                <button
                  onClick={() => handleSelectTab('employee_master')}
                  className={`w-full flex items-center gap-2.5 px-3 py-2 rounded-xl text-xs font-medium ${
                    activeTab === 'employee_master'
                      ? 'bg-[#051C2C] text-white font-semibold'
                      : 'text-[#051C2C] hover:bg-gray-100'
                  }`}
                >
                  <Users className="w-4 h-4" />
                  <span>Employees</span>
                </button>

                {/* Department Sheets */}
                <div className="pt-2">
                  <div className="px-2 py-1 text-[10px] font-bold text-[#888888] uppercase tracking-wider flex items-center gap-1">
                    <Building2 className="w-3 h-3 text-[#2251FF]" />
                    <span>Department Sheets</span>
                  </div>
                  <div className="space-y-1 mt-1 pl-2 border-l-2 border-[#E8E8E6] ml-2">
                    {DEPARTMENTS.map((dept) => {
                      const isSelected = activeTab === dept.id;
                      return (
                        <button
                          key={dept.id}
                          onClick={() => handleSelectTab(dept.id as NavigationTab)}
                          className={`w-full text-left px-2.5 py-1.5 rounded-lg text-xs flex items-center justify-between ${
                            isSelected
                              ? 'bg-[#051C2C] text-white font-semibold'
                              : 'text-[#051C2C] hover:bg-gray-100'
                          }`}
                        >
                          <span className="truncate">{dept.shortName}: {dept.name}</span>
                          <span
                            className={`text-[9px] font-mono font-bold px-1.5 py-0.2 rounded ${
                              isSelected ? 'bg-white/20 text-white' : 'bg-gray-200 text-[#051C2C]'
                            }`}
                          >
                            {dept.deptCode}
                          </span>
                        </button>
                      );
                    })}
                  </div>
                </div>

                <button
                  onClick={() => handleSelectTab('payroll_summary')}
                  className={`w-full flex items-center gap-2.5 px-3 py-2 rounded-xl text-xs font-medium ${
                    activeTab === 'payroll_summary'
                      ? 'bg-[#051C2C] text-white font-semibold'
                      : 'text-[#051C2C] hover:bg-gray-100'
                  }`}
                >
                  <PieChart className="w-4 h-4" />
                  <span>Payroll Summary</span>
                </button>

                <button
                  onClick={() => handleSelectTab('annual_hours_tracker')}
                  className={`w-full flex items-center gap-2.5 px-3 py-2 rounded-xl text-xs font-medium ${
                    activeTab === 'annual_hours_tracker'
                      ? 'bg-[#051C2C] text-white font-semibold'
                      : 'text-[#051C2C] hover:bg-gray-100'
                  }`}
                >
                  <Clock className="w-4 h-4" />
                  <span>Hours Tracker</span>
                </button>
              </div>

              {/* Data Tools */}
              <div className="pt-3 border-t border-[#E8E8E6] space-y-1.5">
                <div className="px-2 py-1 text-[10px] font-bold text-[#888888] uppercase tracking-wider">
                  Data Persistence
                </div>
                <div className="grid grid-cols-2 gap-2">
                  <button
                    onClick={() => {
                      setIsMobileMenuOpen(false);
                      onExportBackup();
                    }}
                    className="p-2 rounded-lg bg-[#F5F5F2] hover:bg-gray-200 text-[#051C2C] text-xs font-semibold flex items-center justify-center gap-1 border border-[#E8E8E6]"
                  >
                    <Download className="w-3.5 h-3.5 text-[#2251FF]" />
                    <span>Export</span>
                  </button>
                  <button
                    onClick={() => {
                      setIsMobileMenuOpen(false);
                      onOpenImportBackup();
                    }}
                    className="p-2 rounded-lg bg-[#F5F5F2] hover:bg-gray-200 text-[#051C2C] text-xs font-semibold flex items-center justify-center gap-1 border border-[#E8E8E6]"
                  >
                    <Upload className="w-3.5 h-3.5 text-[#2251FF]" />
                    <span>Import</span>
                  </button>
                </div>
                <div className="grid grid-cols-2 gap-2">
                  <button
                    onClick={() => {
                      setIsMobileMenuOpen(false);
                      onOpenBulkCsv();
                    }}
                    className="p-2 rounded-lg bg-[#F5F5F2] hover:bg-gray-200 text-[#051C2C] text-xs font-semibold flex items-center justify-center gap-1 border border-[#E8E8E6]"
                  >
                    <FileSpreadsheet className="w-3.5 h-3.5 text-[#00C853]" />
                    <span>Bulk CSV</span>
                  </button>
                  <button
                    onClick={() => {
                      setIsMobileMenuOpen(false);
                      onOpenResetModal();
                    }}
                    className="p-2 rounded-lg bg-red-50 hover:bg-red-100 text-[#D32F2F] text-xs font-semibold flex items-center justify-center gap-1 border border-red-200"
                  >
                    <RotateCcw className="w-3.5 h-3.5 text-[#D32F2F]" />
                    <span>Reset</span>
                  </button>
                </div>
              </div>
            </div>

            {/* Drawer Footer Notice */}
            <div className="pt-3 border-t border-[#E8E8E6] space-y-1 text-[10px] text-[#888888]">
              <div className="flex items-center gap-1 font-semibold text-[#051C2C]">
                <ShieldCheck className="w-3.5 h-3.5 text-[#2251FF]" />
                <span>Local Storage Persistence</span>
              </div>
              <p className="leading-tight">
                All data in this tool is stored in your browser's localStorage. The page itself does not retain any user data.
              </p>
            </div>
          </div>
          <div className="flex-1" onClick={() => setIsMobileMenuOpen(false)} />
        </div>
      )}
    </>
  );
};
