import { Link } from 'react-router-dom';
import React, { useState, useRef } from 'react';
import * as XLSX from 'xlsx';

type EtaEntry = {
  id: number;
  srNo: number;
  vesselName: string;
  port: string;
  etaEtbEtd: string;
  remarks: string;
  portGroup: string;
};

// ----- CONFIGURATION (matching VBA reference) -----
const VALID_CODES = ['YA', 'AF', 'DA', 'BH']; // column G filter
const SORTED_GROUPS = [
  'Fujairah / East Coast',
  'Dubai / Jebel Ali Area / Sharjah / Hamriyah',
  'Abu Dhabi Area',
];
const OTHER_GROUP = 'Other / Unknown';

// Port grouping function (exactly like VBA GetPortGroup)
function getPortGroup(portName: string): string {
  if (!portName) return OTHER_GROUP;
  const port = portName.toString().trim().toUpperCase();

  // Group 1: Fujairah / East Coast
  if (
    port === 'FUJAIRAH' ||
    port === 'DIBBA' ||
    port === 'KHOR FAKKAN' ||
    port === 'KFK' ||
    port === 'MINA SAQR' ||
    port === 'RAS AL KHAIMAH'
  ) {
    return 'Fujairah / East Coast';
  }

  // Group 2: Dubai / Jebel Ali Area / Sharjah / Hamriyah
  if (
    port === 'DUBAI' ||
    port === 'JEBEL ALI' ||
    port === 'JEBAL ALI' ||
    port === 'DUBAI MARITIME CITY' ||
    port === 'DMC' ||
    port === 'SHARJAH' ||
    port === 'HAMRIYAH SHARJAH'
  ) {
    return 'Dubai / Jebel Ali Area / Sharjah / Hamriyah';
  }

  // Group 3: Abu Dhabi Area
  if (
    port === 'ABU DHABI' ||
    port === 'ABU DHABI PORT' ||
    port === 'KHALIFA PORT' ||
    port === 'KHALID PORT'
  ) {
    return 'Abu Dhabi Area';
  }

  return OTHER_GROUP;
}

export default function EtaInfo() {
  const [data, setData] = useState<EtaEntry[]>([]);
  const [searchTerm, setSearchTerm] = useState('');
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [editingId, setEditingId] = useState<number | null>(null);
  const [notification, setNotification] = useState<{
    type: 'success' | 'error';
    message: string;
  } | null>(null);
  const [viewMode, setViewMode] = useState<'flat' | 'grouped'>('flat');
  const [importPreview, setImportPreview] = useState<{
    vesselMap: Map<string, { port: string; group: string }>;
    fileName: string;
  } | null>(null);

  const fileInputRef = useRef<HTMLInputElement>(null);

  // Form state
  const [formData, setFormData] = useState({
    vesselName: '',
    port: '',
    etaEtbEtd: '',
    remarks: '',
  });

  const filteredData = data.filter(
    (item) =>
      item.vesselName.toLowerCase().includes(searchTerm.toLowerCase()) ||
      item.port.toLowerCase().includes(searchTerm.toLowerCase()) ||
      item.remarks.toLowerCase().includes(searchTerm.toLowerCase())
  );

  // Group data by port group for grouped view
  const groupedData = data.reduce<Record<string, EtaEntry[]>>((acc, item) => {
    const group = item.portGroup;
    if (!acc[group]) acc[group] = [];
    acc[group].push(item);
    return acc;
  }, {});

  const showNotification = (type: 'success' | 'error', message: string) => {
    setNotification({ type, message });
    setTimeout(() => setNotification(null), 4000);
  };

  const resetForm = () => {
    setFormData({ vesselName: '', port: '', etaEtbEtd: '', remarks: '' });
    setEditingId(null);
  };

  const openAddModal = () => {
    resetForm();
    setIsModalOpen(true);
  };

  const openEditModal = (item: EtaEntry) => {
    setFormData({
      vesselName: item.vesselName,
      port: item.port,
      etaEtbEtd: item.etaEtbEtd,
      remarks: item.remarks,
    });
    setEditingId(item.id);
    setIsModalOpen(true);
  };

  const closeModal = () => {
    setIsModalOpen(false);
    resetForm();
  };

  const handleFormSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (
      !formData.vesselName.trim() ||
      !formData.port.trim() ||
      !formData.etaEtbEtd.trim()
    ) {
      showNotification(
        'error',
        'Please fill in all required fields (Vessel Name, Port, ETA-ETB-ETD).'
      );
      return;
    }

    if (editingId !== null) {
      setData(
        data.map((item) =>
          item.id === editingId
            ? {
                ...item,
                ...formData,
                portGroup: getPortGroup(formData.port),
              }
            : item
        )
      );
      showNotification('success', 'Entry updated successfully!');
    } else {
      const newEntry: EtaEntry = {
        id: Date.now(),
        srNo: data.length > 0 ? Math.max(...data.map((d) => d.srNo)) + 1 : 1,
        ...formData,
        portGroup: getPortGroup(formData.port),
      };
      setData([...data, newEntry]);
      showNotification('success', 'New entry added successfully!');
    }
    closeModal();
  };

  const handleDelete = (id: number) => {
    if (window.confirm('Are you sure you want to delete this entry?')) {
      const updatedData = data.filter((item) => item.id !== id);
      const reNumbered = updatedData.map((item, idx) => ({
        ...item,
        srNo: idx + 1,
      }));
      setData(reNumbered);
      showNotification('success', 'Entry deleted successfully!');
    }
  };

  // ========== IMPORT FROM EXCEL (matching VBA reference logic) ==========
  const handleImportClick = () => {
    fileInputRef.current?.click();
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();

    reader.onload = (evt) => {
      try {
        const arrayBuffer = evt.target?.result as ArrayBuffer;
        const data = new Uint8Array(arrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });

        // Use the first sheet (like ActiveSheet in VBA)
        const sheetName = workbook.SheetNames[0];
        if (!sheetName) {
          showNotification('error', 'The workbook contains no sheets.');
          return;
        }
        const sheet = workbook.Sheets[sheetName];

        // Convert sheet to array of arrays
        const rows = XLSX.utils.sheet_to_json(sheet, {
          header: 1,
          defval: '',
        }) as any[][];

        if (rows.length < 2) {
          showNotification(
            'error',
            'No data found in the source sheet (less than 2 rows).'
          );
          return;
        }

        // VBA uses columns D (index 3), G (index 6), K (index 10) — 0-based
        const COL_D = 3; // vessel name
        const COL_G = 6; // code
        const COL_K = 10; // port

        // Dictionary to track unique vessels: Map vesselName -> { port, group }
        const vesselMap = new Map<string, { port: string; group: string }>();

        // Start from row index 1 (skip header row like VBA starts at i = 2)
        for (let i = 1; i < rows.length; i++) {
          const row = rows[i];
          if (!row || row.length === 0) continue;

          // Get column G value (code) — trim & uppercase
          const rawCode =
            row[COL_G] !== undefined
              ? row[COL_G].toString().trim().toUpperCase()
              : '';
          const isValidCode = VALID_CODES.includes(rawCode);
          if (!isValidCode) continue;

          // Get vessel name from column D
          const rawVessel =
            row[COL_D] !== undefined ? row[COL_D].toString().trim() : '';
          if (rawVessel === '') continue;

          // Only add if not already in map (unique vessels)
          if (!vesselMap.has(rawVessel)) {
            // Get port from column K
            const rawPort =
              row[COL_K] !== undefined
                ? row[COL_K].toString().trim().toUpperCase()
                : '';
            const portGroup = getPortGroup(rawPort);
            vesselMap.set(rawVessel, { port: rawPort, group: portGroup });
          }
        }

        if (vesselMap.size === 0) {
          showNotification(
            'error',
            'No data found matching the criteria (YA, AF, DA, BH).'
          );
          return;
        }

        // Show preview before importing
        setImportPreview({
          vesselMap,
          fileName: file.name,
        });
      } catch (err) {
        console.error(err);
        showNotification(
          'error',
          'Failed to read Excel file. Please check the format.'
        );
      }
    };

    reader.onerror = () => {
      showNotification('error', 'Failed to read the file.');
    };

    reader.readAsArrayBuffer(file);

    // Reset file input so same file can be re-imported
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  // Confirm import — populate the ETA table
  const confirmImport = () => {
    if (!importPreview) return;

    const importedEntries: EtaEntry[] = [];
    let srCounter = 1;

    // Process in group order (matching VBA output)
    SORTED_GROUPS.forEach((groupName) => {
      importPreview.vesselMap.forEach((value, vessel) => {
        if (value.group === groupName) {
          importedEntries.push({
            id: Date.now() + srCounter,
            srNo: srCounter++,
            vesselName: vessel,
            port: value.port,
            etaEtbEtd: '', // User fills in later
            remarks: '', // User fills in later
            portGroup: value.group,
          });
        }
      });
    });

    // Process "Other / Unknown" group
    importPreview.vesselMap.forEach((value, vessel) => {
      if (value.group === OTHER_GROUP) {
        importedEntries.push({
          id: Date.now() + srCounter,
          srNo: srCounter++,
          vesselName: vessel,
          port: value.port,
          etaEtbEtd: '',
          remarks: '',
          portGroup: value.group,
        });
      }
    });

    setData(importedEntries);
    showNotification(
      'success',
      `Successfully imported ${importedEntries.length} vessels! You can now edit each row to add ETA/ETB/ETD and remarks.`
    );
    setImportPreview(null);
  };

  const cancelImport = () => {
    setImportPreview(null);
  };

  const downloadTemplate = () => {
    const templateData = [
      ['SR#', 'VESSEL NAME', 'PORT', 'ETA - ETB - ETD', 'REMARKS'],
      [
        1,
        'MSC ARIES',
        'Jebel Ali, UAE',
        'ETA: 15 Nov / ETB: 16 Nov / ETD: 18 Nov',
        'FFV confirmed',
      ],
    ];
    const ws = XLSX.utils.aoa_to_sheet(templateData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'ETA Template');
    XLSX.writeFile(wb, 'ETA_Import_Template.xlsx');
  };

  // ========== SEND BUTTON (placeholder — code to be provided) ==========
  const handleSend = () => {
    if (data.length === 0) {
      showNotification('error', 'No data to send. Please add entries or import from Excel first.');
      return;
    }
    // TODO: Replace this with the actual send logic once code is provided
    showNotification('success', `Send button clicked! ${data.length} entries ready. Awaiting send logic implementation.`);
  };

  // Build preview grouped list
  const buildPreviewGroups = () => {
    if (!importPreview) return [];
    const groups: { name: string; vessels: { vessel: string; port: string }[] }[] = [];

    SORTED_GROUPS.forEach((groupName) => {
      const vessels: { vessel: string; port: string }[] = [];
      importPreview.vesselMap.forEach((value, vessel) => {
        if (value.group === groupName) {
          vessels.push({ vessel, port: value.port });
        }
      });
      if (vessels.length > 0) {
        groups.push({ name: groupName, vessels });
      }
    });

    const otherVessels: { vessel: string; port: string }[] = [];
    importPreview.vesselMap.forEach((value, vessel) => {
      if (value.group === OTHER_GROUP) {
        otherVessels.push({ vessel, port: value.port });
      }
    });
    if (otherVessels.length > 0) {
      groups.push({ name: OTHER_GROUP, vessels: otherVessels });
    }

    return groups;
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-blue-50">
      {/* Notification Toast */}
      {notification && (
        <div className="fixed top-6 right-6 z-[60] animate-slide-in">
          <div
            className={`flex items-center gap-3 px-5 py-3 rounded-lg shadow-xl border max-w-md ${
              notification.type === 'success'
                ? 'bg-emerald-50 border-emerald-200 text-emerald-800'
                : 'bg-red-50 border-red-200 text-red-800'
            }`}
          >
            <i
              className={`fa-solid ${
                notification.type === 'success'
                  ? 'fa-circle-check'
                  : 'fa-circle-exclamation'
              }`}
            ></i>
            <span className="text-sm font-medium">{notification.message}</span>
          </div>
        </div>
      )}

      {/* Hidden file input */}
      <input
        type="file"
        ref={fileInputRef}
        onChange={handleFileChange}
        accept=".xlsx,.xls,.xlsm,.xlsb,.csv"
        className="hidden"
      />

      {/* Navigation Bar */}
      <nav className="bg-[#0B1D3A] shadow-lg">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex items-center justify-between h-16">
            <Link to="/" className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-full bg-white/10 flex items-center justify-center overflow-hidden border border-[#D4A843]/30">
                <img
                  src="https://lh3.googleusercontent.com/d/1WBc2Su-36a-le49VAP9rS4pRCOjO1qxw"
                  alt="Ocean Fair Logo"
                  className="w-9 h-9 object-contain"
                />
              </div>
              <span className="text-white font-bold text-sm hidden sm:block">
                Ocean Fair International Group FZE
              </span>
            </Link>
            <Link
              to="/"
              className="text-blue-300 hover:text-white text-sm flex items-center gap-2 transition-colors"
            >
              <i className="fa-solid fa-arrow-left"></i> Back to Dashboard
            </Link>
          </div>
        </div>
      </nav>

      {/* Page Header */}
      <section className="bg-gradient-to-r from-[#0B1D3A] via-[#1A3A6B] to-[#0B1D3A]">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-10">
          <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-4">
            <div className="flex items-center gap-4">
              <div className="w-14 h-14 rounded-xl bg-blue-500/20 backdrop-blur-sm flex items-center justify-center border border-blue-400/30">
                <i className="fa-solid fa-ship text-blue-300 text-2xl"></i>
              </div>
              <div>
                <h1 className="text-3xl font-bold text-white">ETA INFO</h1>
                <p className="text-blue-200 text-sm mt-1">
                  Track vessel arrivals, berthing & departure schedules
                </p>
              </div>
            </div>
            <div className="flex flex-wrap items-center gap-3">
              <button
                onClick={handleSend}
                className="bg-gradient-to-r from-indigo-600 to-blue-600 hover:from-indigo-700 hover:to-blue-700 text-white font-semibold px-5 py-3 rounded-lg shadow-lg transition-all flex items-center gap-2"
              >
                <i className="fa-solid fa-paper-plane"></i> Send
              </button>
              <button
                onClick={handleImportClick}
                className="bg-emerald-600 hover:bg-emerald-700 text-white font-semibold px-5 py-3 rounded-lg shadow-lg transition-all flex items-center gap-2"
              >
                <i className="fa-solid fa-file-excel"></i> Import from Excel
              </button>
              <button
                onClick={downloadTemplate}
                className="bg-white/10 hover:bg-white/20 backdrop-blur-sm text-white font-semibold px-5 py-3 rounded-lg border border-white/20 transition-all flex items-center gap-2"
              >
                <i className="fa-solid fa-download"></i> Template
              </button>
              <button
                onClick={openAddModal}
                className="bg-[#D4A843] hover:bg-[#b8912f] text-white font-semibold px-6 py-3 rounded-lg shadow-lg transition-all flex items-center gap-2"
              >
                <i className="fa-solid fa-plus"></i> Add New Entry
              </button>
            </div>
          </div>
        </div>
      </section>

      {/* Main Content */}
      <section className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        {/* Search & View Toggle Bar */}
        <div className="bg-white rounded-xl shadow-md p-4 mb-6 flex flex-col md:flex-row gap-4 items-center justify-between">
          <div className="relative flex-1 w-full md:max-w-md">
            <i className="fa-solid fa-search absolute left-4 top-1/2 -translate-y-1/2 text-gray-400"></i>
            <input
              type="text"
              placeholder="Search by vessel name, port, or remarks..."
              value={searchTerm}
              onChange={(e) => setSearchTerm(e.target.value)}
              className="w-full pl-11 pr-4 py-2.5 border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-[#1A3A6B] focus:border-transparent text-sm"
            />
          </div>
          <div className="flex items-center gap-3">
            <span className="text-sm text-gray-500">
              <span className="font-semibold text-[#0B1D3A]">
                {filteredData.length}
              </span>{' '}
              entries found
            </span>
            {data.length > 0 && (
              <div className="flex items-center bg-gray-100 rounded-lg p-1">
                <button
                  onClick={() => setViewMode('flat')}
                  className={`px-3 py-1.5 text-xs font-medium rounded-md transition-colors ${
                    viewMode === 'flat'
                      ? 'bg-white text-[#0B1D3A] shadow-sm'
                      : 'text-gray-500 hover:text-gray-700'
                  }`}
                >
                  <i className="fa-solid fa-list mr-1"></i> Flat
                </button>
                <button
                  onClick={() => setViewMode('grouped')}
                  className={`px-3 py-1.5 text-xs font-medium rounded-md transition-colors ${
                    viewMode === 'grouped'
                      ? 'bg-white text-[#0B1D3A] shadow-sm'
                      : 'text-gray-500 hover:text-gray-700'
                  }`}
                >
                  <i className="fa-solid fa-layer-group mr-1"></i> Grouped
                </button>
              </div>
            )}
          </div>
        </div>

        {/* Info Banner — Column mapping */}
        <div className="mb-6 bg-blue-50 border border-blue-200 rounded-xl p-4 flex items-start gap-3">
          <div className="w-10 h-10 rounded-lg bg-blue-100 flex items-center justify-center flex-shrink-0">
            <i className="fa-solid fa-circle-info text-blue-600"></i>
          </div>
          <div className="text-sm text-blue-900">
            <p className="font-semibold mb-1">Excel Import — Column Mapping</p>
            <p className="text-blue-800">
              Columns used from source file:{' '}
              <strong>D</strong> (Vessel Name) •{' '}
              <strong>G</strong> (Code: YA, AF, DA, BH) •{' '}
              <strong>K</strong> (Port). Vessels are grouped by port region
              (Fujairah/East Coast, Dubai/Jebel Ali/Sharjah/Hamriyah, Abu Dhabi).
            </p>
          </div>
        </div>

        {/* FLAT TABLE VIEW */}
        {viewMode === 'flat' && (
          <div className="bg-white rounded-xl shadow-md overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full">
                <thead>
                  <tr className="bg-[#0B1D3A] text-white">
                    <th className="px-4 py-4 text-left text-xs font-bold uppercase tracking-wider border-r border-white/10">
                      SR#
                    </th>
                    <th className="px-4 py-4 text-left text-xs font-bold uppercase tracking-wider border-r border-white/10">
                      VESSEL NAME
                    </th>
                    <th className="px-4 py-4 text-left text-xs font-bold uppercase tracking-wider border-r border-white/10">
                      PORT
                    </th>
                    <th className="px-4 py-4 text-center text-xs font-bold uppercase tracking-wider border-r border-white/10">
                      ETA - ETB - ETD
                    </th>
                    <th className="px-4 py-4 text-left text-xs font-bold uppercase tracking-wider border-r border-white/10">
                      REMARKS (Confirm for FFV, Bread &amp; Dairy arrangement)
                    </th>
                    <th className="px-4 py-4 text-center text-xs font-bold uppercase tracking-wider">
                      ACTION
                    </th>
                  </tr>
                </thead>

                <tbody className="divide-y divide-gray-100">
                  {filteredData.length === 0 ? (
                    <tr>
                      <td
                        colSpan={6}
                        className="px-4 py-16 text-center text-gray-400"
                      >
                        <i className="fa-solid fa-inbox text-5xl mb-4 block text-gray-300"></i>
                        <p className="text-base font-medium text-gray-500 mb-1">
                          No entries yet
                        </p>
                        <p className="text-sm">
                          Click{' '}
                          <span className="font-semibold text-[#D4A843]">
                            "Add New Entry"
                          </span>{' '}
                          or{' '}
                          <span className="font-semibold text-emerald-600">
                            "Import from Excel"
                          </span>{' '}
                          to get started.
                        </p>
                      </td>
                    </tr>
                  ) : (
                    filteredData.map((item, index) => (
                      <tr
                        key={item.id}
                        className={`${
                          index % 2 === 0 ? 'bg-white' : 'bg-slate-50'
                        } hover:bg-blue-50 transition-colors`}
                      >
                        <td className="px-4 py-4 text-sm font-semibold text-[#0B1D3A] border-r border-gray-100">
                          {item.srNo}
                        </td>
                        <td className="px-4 py-4 text-sm font-semibold text-gray-800 border-r border-gray-100">
                          <div className="flex items-center gap-2">
                            <i className="fa-solid fa-ship text-blue-400 text-xs"></i>
                            {item.vesselName}
                          </div>
                        </td>
                        <td className="px-4 py-4 text-sm text-gray-700 border-r border-gray-100">
                          <div className="flex items-center gap-2">
                            <i className="fa-solid fa-location-dot text-[#D4A843] text-xs"></i>
                            {item.port}
                          </div>
                        </td>
                        <td className="px-4 py-4 text-sm text-center text-gray-700 border-r border-gray-100">
                          {item.etaEtbEtd || (
                            <span className="text-gray-300 italic text-xs">
                              — not set —
                            </span>
                          )}
                        </td>
                        <td className="px-4 py-4 text-sm text-gray-600 border-r border-gray-100 max-w-xs">
                          <span className="inline-block">
                            {item.remarks || (
                              <span className="text-gray-300 italic text-xs">
                                — not set —
                              </span>
                            )}
                          </span>
                        </td>
                        <td className="px-4 py-4 text-center">
                          <div className="flex items-center justify-center gap-2">
                            <button
                              onClick={() => openEditModal(item)}
                              className="w-8 h-8 rounded-lg bg-amber-50 hover:bg-amber-100 text-amber-600 flex items-center justify-center transition-colors"
                              title="Edit"
                            >
                              <i className="fa-solid fa-pen-to-square text-xs"></i>
                            </button>
                            <button
                              onClick={() => handleDelete(item.id)}
                              className="w-8 h-8 rounded-lg bg-red-50 hover:bg-red-100 text-red-600 flex items-center justify-center transition-colors"
                              title="Delete"
                            >
                              <i className="fa-solid fa-trash text-xs"></i>
                            </button>
                          </div>
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>

            {/* Table Footer */}
            <div className="bg-slate-50 px-4 py-3 border-t border-gray-200 flex items-center justify-between">
              <p className="text-sm text-gray-500">
                Showing{' '}
                <span className="font-semibold text-[#0B1D3A]">
                  {filteredData.length}
                </span>{' '}
                of{' '}
                <span className="font-semibold text-[#0B1D3A]">
                  {data.length}
                </span>{' '}
                entries
              </p>
            </div>
          </div>
        )}

        {/* GROUPED TABLE VIEW */}
        {viewMode === 'grouped' && data.length > 0 && (
          <div className="bg-white rounded-xl shadow-md overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full">
                <thead>
                  <tr className="bg-[#0B1D3A] text-white">
                    <th className="px-4 py-4 text-left text-xs font-bold uppercase tracking-wider">
                      SHIP TO / VESSEL
                    </th>
                    <th className="px-4 py-4 text-left text-xs font-bold uppercase tracking-wider">
                      DELIVERY PORT
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {buildPreviewGroups().flatMap((group, gIdx) => {
                    const rows: React.ReactNode[] = [];
                    // Group header row
                    rows.push(
                      <tr key={`group-${gIdx}`} className="bg-slate-100 border-y-2 border-slate-300">
                        <td
                          colSpan={2}
                          className="px-5 py-3 font-bold text-[#0B1D3A] text-sm tracking-wide"
                        >
                          <i className="fa-solid fa-layer-group text-[#D4A843] mr-2"></i>
                          {group.name}
                          <span className="ml-2 text-xs font-normal text-gray-500">
                            ({group.vessels.length} vessel
                            {group.vessels.length !== 1 ? 's' : ''})
                          </span>
                        </td>
                      </tr>
                    );
                    // Vessel rows
                    group.vessels.forEach((v, vIdx) => {
                      rows.push(
                        <tr
                          key={`vessel-${gIdx}-${vIdx}`}
                          className="hover:bg-blue-50 transition-colors border-b border-gray-100"
                        >
                          <td className="px-5 py-3 text-sm font-medium text-[#0B1D3A]">
                            <i className="fa-solid fa-ship text-blue-400 text-xs mr-2"></i>
                            {v.vessel}
                          </td>
                          <td className="px-5 py-3 text-sm text-gray-600">
                            <i className="fa-solid fa-location-dot text-[#D4A843] text-xs mr-2"></i>
                            {v.port}
                          </td>
                        </tr>
                      );
                    });
                    return rows;
                  })}
                </tbody>
              </table>
            </div>
            <div className="bg-slate-50 px-4 py-3 border-t border-gray-200">
              <p className="text-sm text-gray-500">
                Total:{' '}
                <span className="font-semibold text-[#0B1D3A]">
                  {data.length}
                </span>{' '}
                vessels across{' '}
                <span className="font-semibold text-[#0B1D3A]">
                  {buildPreviewGroups().length}
                </span>{' '}
                port region(s)
              </p>
            </div>
          </div>
        )}

        {/* Important Notice */}
        <div className="mt-6 bg-amber-50 border border-amber-200 rounded-xl p-4 flex items-start gap-3">
          <div className="w-10 h-10 rounded-lg bg-amber-100 flex items-center justify-center flex-shrink-0">
            <i className="fa-solid fa-circle-info text-amber-600"></i>
          </div>
          <div>
            <h4 className="font-semibold text-amber-900 text-sm mb-1">
              Important Notice
            </h4>
            <p className="text-amber-800 text-sm">
              Please confirm all arrangements for{' '}
              <strong>FFV (Fresh Frozen Vegetables)</strong>,{' '}
              <strong>Bread</strong>, and <strong>Dairy</strong> products prior
              to vessel arrival to ensure smooth clearance and delivery.
            </p>
          </div>
        </div>
      </section>

      {/* IMPORT PREVIEW MODAL */}
      {importPreview && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-2xl max-h-[90vh] overflow-hidden flex flex-col">
            {/* Modal Header */}
            <div className="bg-gradient-to-r from-emerald-700 to-emerald-900 px-6 py-4 flex items-center justify-between flex-shrink-0">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-lg bg-white/10 flex items-center justify-center">
                  <i className="fa-solid fa-file-excel text-white"></i>
                </div>
                <div>
                  <h3 className="text-white font-bold text-lg">
                    Import Preview
                  </h3>
                  <p className="text-emerald-100 text-xs">
                    {importPreview.fileName}
                  </p>
                </div>
              </div>
              <button
                onClick={cancelImport}
                className="text-white/70 hover:text-white transition-colors"
              >
                <i className="fa-solid fa-xmark text-xl"></i>
              </button>
            </div>

            {/* Modal Body */}
            <div className="p-6 overflow-y-auto flex-1">
              <div className="mb-4 bg-emerald-50 border border-emerald-200 rounded-lg p-3 flex items-center gap-2">
                <i className="fa-solid fa-circle-check text-emerald-600"></i>
                <span className="text-sm text-emerald-800 font-medium">
                  {importPreview.vesselMap.size} unique vessel
                  {importPreview.vesselMap.size !== 1 ? 's' : ''} extracted
                  (filtered by codes: YA, AF, DA, BH)
                </span>
              </div>

              <p className="text-sm text-gray-600 mb-4">
                Vessels will be grouped by port region. ETA/ETB/ETD and Remarks
                fields will be empty — you can edit each row later to fill them
                in.
              </p>

              <div className="space-y-3 max-h-96 overflow-y-auto">
                {buildPreviewGroups().map((group, gIdx) => (
                  <div
                    key={gIdx}
                    className="border border-gray-200 rounded-lg overflow-hidden"
                  >
                    <div className="bg-slate-100 px-4 py-2 border-b border-gray-200">
                      <p className="text-sm font-bold text-[#0B1D3A]">
                        <i className="fa-solid fa-layer-group text-[#D4A843] mr-2"></i>
                        {group.name}
                        <span className="ml-2 text-xs font-normal text-gray-500">
                          ({group.vessels.length})
                        </span>
                      </p>
                    </div>
                    <div className="divide-y divide-gray-100">
                      {group.vessels.map((v, vIdx) => (
                        <div
                          key={vIdx}
                          className="px-4 py-2 flex items-center justify-between text-sm hover:bg-blue-50"
                        >
                          <span className="font-medium text-gray-800 flex items-center gap-2">
                            <i className="fa-solid fa-ship text-blue-400 text-xs"></i>
                            {v.vessel}
                          </span>
                          <span className="text-gray-500 text-xs flex items-center gap-1">
                            <i className="fa-solid fa-location-dot text-[#D4A843]"></i>
                            {v.port}
                          </span>
                        </div>
                      ))}
                    </div>
                  </div>
                ))}
              </div>
            </div>

            {/* Modal Footer */}
            <div className="border-t border-gray-100 px-6 py-4 flex items-center justify-end gap-3 bg-gray-50 flex-shrink-0">
              <button
                onClick={cancelImport}
                className="px-5 py-2.5 text-sm font-medium text-gray-700 bg-white border border-gray-200 hover:bg-gray-50 rounded-lg transition-colors"
              >
                Cancel
              </button>
              <button
                onClick={confirmImport}
                className="px-5 py-2.5 text-sm font-semibold text-white bg-emerald-600 hover:bg-emerald-700 rounded-lg transition-colors flex items-center gap-2"
              >
                <i className="fa-solid fa-check"></i>
                Confirm Import ({importPreview.vesselMap.size} vessels)
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ADD/EDIT ENTRY MODAL */}
      {isModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-lg overflow-hidden">
            {/* Modal Header */}
            <div className="bg-gradient-to-r from-[#0B1D3A] to-[#1A3A6B] px-6 py-4 flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-lg bg-white/10 flex items-center justify-center">
                  <i
                    className={`fa-solid ${
                      editingId ? 'fa-pen-to-square' : 'fa-plus'
                    } text-[#D4A843]`}
                  ></i>
                </div>
                <h3 className="text-white font-bold text-lg">
                  {editingId ? 'Edit Entry' : 'Add New Entry'}
                </h3>
              </div>
              <button
                onClick={closeModal}
                className="text-white/70 hover:text-white transition-colors"
              >
                <i className="fa-solid fa-xmark text-xl"></i>
              </button>
            </div>

            {/* Modal Body */}
            <form onSubmit={handleFormSubmit} className="p-6 space-y-4">
              <div>
                <label className="block text-sm font-semibold text-gray-700 mb-1.5">
                  Vessel Name <span className="text-red-500">*</span>
                </label>
                <div className="relative">
                  <i className="fa-solid fa-ship absolute left-3 top-1/2 -translate-y-1/2 text-gray-400 text-sm"></i>
                  <input
                    type="text"
                    value={formData.vesselName}
                    onChange={(e) =>
                      setFormData({ ...formData, vesselName: e.target.value })
                    }
                    placeholder="e.g. MSC ARIES"
                    className="w-full pl-10 pr-4 py-2.5 border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-[#1A3A6B] focus:border-transparent text-sm"
                    required
                  />
                </div>
              </div>

              <div>
                <label className="block text-sm font-semibold text-gray-700 mb-1.5">
                  Port <span className="text-red-500">*</span>
                </label>
                <div className="relative">
                  <i className="fa-solid fa-location-dot absolute left-3 top-1/2 -translate-y-1/2 text-gray-400 text-sm"></i>
                  <input
                    type="text"
                    value={formData.port}
                    onChange={(e) =>
                      setFormData({ ...formData, port: e.target.value })
                    }
                    placeholder="e.g. JEBEL ALI, FUJAIRAH, ABU DHABI"
                    className="w-full pl-10 pr-4 py-2.5 border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-[#1A3A6B] focus:border-transparent text-sm"
                    required
                  />
                </div>
                <p className="text-xs text-gray-400 mt-1">
                  Port will be auto-grouped (Fujairah/East Coast, Dubai/Jebel
                  Ali/Sharjah/Hamriyah, Abu Dhabi, or Other)
                </p>
              </div>

              <div>
                <label className="block text-sm font-semibold text-gray-700 mb-1.5">
                  ETA - ETB - ETD <span className="text-red-500">*</span>
                </label>
                <div className="relative">
                  <i className="fa-regular fa-calendar absolute left-3 top-3 text-gray-400 text-sm"></i>
                  <textarea
                    value={formData.etaEtbEtd}
                    onChange={(e) =>
                      setFormData({ ...formData, etaEtbEtd: e.target.value })
                    }
                    placeholder="e.g. ETA: 15 Nov / ETB: 16 Nov / ETD: 18 Nov"
                    rows={2}
                    className="w-full pl-10 pr-4 py-2.5 border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-[#1A3A6B] focus:border-transparent text-sm resize-none"
                    required
                  />
                </div>
              </div>

              <div>
                <label className="block text-sm font-semibold text-gray-700 mb-1.5">
                  Remarks (Confirm for FFV, Bread & Dairy arrangement)
                </label>
                <div className="relative">
                  <i className="fa-solid fa-note-sticky absolute left-3 top-3 text-gray-400 text-sm"></i>
                  <textarea
                    value={formData.remarks}
                    onChange={(e) =>
                      setFormData({ ...formData, remarks: e.target.value })
                    }
                    placeholder="Enter remarks here..."
                    rows={2}
                    className="w-full pl-10 pr-4 py-2.5 border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-[#1A3A6B] focus:border-transparent text-sm resize-none"
                  />
                </div>
              </div>

              {/* Modal Footer */}
              <div className="flex items-center justify-end gap-3 pt-4 border-t border-gray-100">
                <button
                  type="button"
                  onClick={closeModal}
                  className="px-5 py-2.5 text-sm font-medium text-gray-700 bg-gray-100 hover:bg-gray-200 rounded-lg transition-colors"
                >
                  Cancel
                </button>
                <button
                  type="submit"
                  className="px-5 py-2.5 text-sm font-semibold text-white bg-[#0B1D3A] hover:bg-[#1A3A6B] rounded-lg transition-colors flex items-center gap-2"
                >
                  <i className="fa-solid fa-check"></i>
                  {editingId ? 'Update Entry' : 'Add Entry'}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Footer */}
      <footer className="bg-[#0B1D3A] text-white mt-12">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6">
          <div className="flex flex-col md:flex-row items-center justify-between gap-4">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-full bg-white/10 flex items-center justify-center overflow-hidden border border-[#D4A843]/30">
                <img
                  src="https://lh3.googleusercontent.com/d/1WBc2Su-36a-le49VAP9rS4pRCOjO1qxw"
                  alt="Ocean Fair Logo"
                  className="w-9 h-9 object-contain"
                />
              </div>
              <div>
                <p className="font-semibold text-sm">
                  Ocean Fair International Group FZE
                </p>
                <p className="text-blue-300 text-xs">
                  Procurement Management Portal
                </p>
              </div>
            </div>
            <p className="text-blue-300 text-sm">
              © {new Date().getFullYear()} Ocean Fair International Group FZE.
              All rights reserved.
            </p>
          </div>
        </div>
      </footer>
    </div>
  );
}
