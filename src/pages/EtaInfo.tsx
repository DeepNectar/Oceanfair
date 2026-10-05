import { Link } from 'react-router-dom';
import { useState, useRef } from 'react';
import * as XLSX from 'xlsx';

type EtaEntry = {
  id: number;
  srNo: number;
  vesselName: string;
  port: string;
  etaEtbEtd: string;
  remarks: string;
};

export default function EtaInfo() {
  const [data, setData] = useState<EtaEntry[]>([]);
  const [searchTerm, setSearchTerm] = useState('');
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [editingId, setEditingId] = useState<number | null>(null);
  const [notification, setNotification] = useState<{ type: 'success' | 'error'; message: string } | null>(null);
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

  const showNotification = (type: 'success' | 'error', message: string) => {
    setNotification({ type, message });
    setTimeout(() => setNotification(null), 3500);
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
    if (!formData.vesselName.trim() || !formData.port.trim() || !formData.etaEtbEtd.trim()) {
      showNotification('error', 'Please fill in all required fields (Vessel Name, Port, ETA-ETB-ETD).');
      return;
    }

    if (editingId !== null) {
      // Update existing
      setData(
        data.map((item) =>
          item.id === editingId
            ? { ...item, ...formData }
            : item
        )
      );
      showNotification('success', 'Entry updated successfully!');
    } else {
      // Add new
      const newEntry: EtaEntry = {
        id: Date.now(),
        srNo: data.length > 0 ? Math.max(...data.map((d) => d.srNo)) + 1 : 1,
        ...formData,
      };
      setData([...data, newEntry]);
      showNotification('success', 'New entry added successfully!');
    }
    closeModal();
  };

  const handleDelete = (id: number) => {
    if (window.confirm('Are you sure you want to delete this entry?')) {
      const updatedData = data.filter((item) => item.id !== id);
      // Re-number SR#
      const reNumbered = updatedData.map((item, idx) => ({ ...item, srNo: idx + 1 }));
      setData(reNumbered);
      showNotification('success', 'Entry deleted successfully!');
    }
  };

  // ========== IMPORT FROM EXCEL ==========
  const handleImportClick = () => {
    fileInputRef.current?.click();
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const workbook = XLSX.read(bstr, { type: 'binary' });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];

        if (jsonData.length < 2) {
          showNotification('error', 'Excel file is empty or has no data rows.');
          return;
        }

        // Skip header row (row 0), parse data rows
        const importedEntries: EtaEntry[] = [];
        let srCounter = data.length > 0 ? Math.max(...data.map((d) => d.srNo)) + 1 : 1;

        for (let i = 1; i < jsonData.length; i++) {
          const row = jsonData[i];
          if (!row || row.length === 0) continue;

          const vesselName = String(row[1] || '').trim();
          const port = String(row[2] || '').trim();
          const etaEtbEtd = String(row[3] || '').trim();
          const remarks = String(row[4] || '').trim();

          if (vesselName || port || etaEtbEtd) {
            importedEntries.push({
              id: Date.now() + i,
              srNo: srCounter++,
              vesselName,
              port,
              etaEtbEtd,
              remarks,
            });
          }
        }

        if (importedEntries.length === 0) {
          showNotification('error', 'No valid data found in the Excel file.');
          return;
        }

        setData([...data, ...importedEntries]);
        showNotification('success', `Successfully imported ${importedEntries.length} entries!`);
      } catch (err) {
        console.error(err);
        showNotification('error', 'Failed to read Excel file. Please check the format.');
      }
    };
    reader.onerror = () => {
      showNotification('error', 'Failed to read the file.');
    };
    reader.readAsBinaryString(file);

    // Reset file input so same file can be re-imported
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const downloadTemplate = () => {
    const templateData = [
      ['SR#', 'VESSEL NAME', 'PORT', 'ETA - ETB - ETD', 'REMARKS'],
      [1, 'MSC ARIES', 'Jebel Ali, UAE', 'ETA: 15 Nov / ETB: 16 Nov / ETD: 18 Nov', 'FFV confirmed'],
    ];
    const ws = XLSX.utils.aoa_to_sheet(templateData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'ETA Template');
    XLSX.writeFile(wb, 'ETA_Import_Template.xlsx');
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-blue-50">
      {/* Notification Toast */}
      {notification && (
        <div className="fixed top-6 right-6 z-50 animate-slide-in">
          <div
            className={`flex items-center gap-3 px-5 py-3 rounded-lg shadow-xl border ${
              notification.type === 'success'
                ? 'bg-emerald-50 border-emerald-200 text-emerald-800'
                : 'bg-red-50 border-red-200 text-red-800'
            }`}
          >
            <i
              className={`fa-solid ${
                notification.type === 'success' ? 'fa-circle-check' : 'fa-circle-exclamation'
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
        accept=".xlsx,.xls,.csv"
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
              <span className="text-white font-bold text-sm hidden sm:block">Ocean Fair International Group FZE</span>
            </Link>
            <Link to="/" className="text-blue-300 hover:text-white text-sm flex items-center gap-2 transition-colors">
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
                <p className="text-blue-200 text-sm mt-1">Track vessel arrivals, berthing & departure schedules</p>
              </div>
            </div>
            <div className="flex flex-wrap items-center gap-3">
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
        {/* Search & Filter Bar */}
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
              <span className="font-semibold text-[#0B1D3A]">{filteredData.length}</span> entries found
            </span>
          </div>
        </div>

        {/* Table */}
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
                    <td colSpan={6} className="px-4 py-16 text-center text-gray-400">
                      <i className="fa-solid fa-inbox text-5xl mb-4 block text-gray-300"></i>
                      <p className="text-base font-medium text-gray-500 mb-1">No entries yet</p>
                      <p className="text-sm">
                        Click <span className="font-semibold text-[#D4A843]">"Add New Entry"</span> or{' '}
                        <span className="font-semibold text-emerald-600">"Import from Excel"</span> to get started.
                      </p>
                    </td>
                  </tr>
                ) : (
                  filteredData.map((item, index) => (
                    <tr
                      key={item.id}
                      className={`${index % 2 === 0 ? 'bg-white' : 'bg-slate-50'} hover:bg-blue-50 transition-colors`}
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
                        {item.etaEtbEtd}
                      </td>
                      <td className="px-4 py-4 text-sm text-gray-600 border-r border-gray-100 max-w-xs">
                        <span className="inline-block">{item.remarks}</span>
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
              Showing <span className="font-semibold text-[#0B1D3A]">{filteredData.length}</span> of{' '}
              <span className="font-semibold text-[#0B1D3A]">{data.length}</span> entries
            </p>
          </div>
        </div>

        {/* Info Box */}
        <div className="mt-6 bg-amber-50 border border-amber-200 rounded-xl p-4 flex items-start gap-3">
          <div className="w-10 h-10 rounded-lg bg-amber-100 flex items-center justify-center flex-shrink-0">
            <i className="fa-solid fa-circle-info text-amber-600"></i>
          </div>
          <div>
            <h4 className="font-semibold text-amber-900 text-sm mb-1">Important Notice</h4>
            <p className="text-amber-800 text-sm">
              Please confirm all arrangements for <strong>FFV (Fresh Frozen Vegetables)</strong>, <strong>Bread</strong>, and <strong>Dairy</strong> products prior to vessel arrival to ensure smooth clearance and delivery.
            </p>
          </div>
        </div>
      </section>

      {/* Add/Edit Entry Modal */}
      {isModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/50 backdrop-blur-sm">
          <div className="bg-white rounded-2xl shadow-2xl w-full max-w-lg overflow-hidden">
            {/* Modal Header */}
            <div className="bg-gradient-to-r from-[#0B1D3A] to-[#1A3A6B] px-6 py-4 flex items-center justify-between">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 rounded-lg bg-white/10 flex items-center justify-center">
                  <i className={`fa-solid ${editingId ? 'fa-pen-to-square' : 'fa-plus'} text-[#D4A843]`}></i>
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
                    onChange={(e) => setFormData({ ...formData, vesselName: e.target.value })}
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
                    onChange={(e) => setFormData({ ...formData, port: e.target.value })}
                    placeholder="e.g. Jebel Ali, UAE"
                    className="w-full pl-10 pr-4 py-2.5 border border-gray-200 rounded-lg focus:outline-none focus:ring-2 focus:ring-[#1A3A6B] focus:border-transparent text-sm"
                    required
                  />
                </div>
              </div>

              <div>
                <label className="block text-sm font-semibold text-gray-700 mb-1.5">
                  ETA - ETB - ETD <span className="text-red-500">*</span>
                </label>
                <div className="relative">
                  <i className="fa-regular fa-calendar absolute left-3 top-3 text-gray-400 text-sm"></i>
                  <textarea
                    value={formData.etaEtbEtd}
                    onChange={(e) => setFormData({ ...formData, etaEtbEtd: e.target.value })}
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
                    onChange={(e) => setFormData({ ...formData, remarks: e.target.value })}
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
                <p className="font-semibold text-sm">Ocean Fair International Group FZE</p>
                <p className="text-blue-300 text-xs">Procurement Management Portal</p>
              </div>
            </div>
            <p className="text-blue-300 text-sm">© {new Date().getFullYear()} Ocean Fair International Group FZE. All rights reserved.</p>
          </div>
        </div>
      </footer>
    </div>
  );
}
