import { Link } from 'react-router-dom';
import { useState } from 'react';

type EtaEntry = {
  id: number;
  srNo: number;
  vesselName: string;
  port: string;
  etaEtbEtd: string;
  remarks: string;
};

const initialData: EtaEntry[] = [];

export default function EtaInfo() {
  const [data, setData] = useState<EtaEntry[]>(initialData);
  const [searchTerm, setSearchTerm] = useState('');

  const filteredData = data.filter(
    (item) =>
      item.vesselName.toLowerCase().includes(searchTerm.toLowerCase()) ||
      item.port.toLowerCase().includes(searchTerm.toLowerCase()) ||
      item.remarks.toLowerCase().includes(searchTerm.toLowerCase())
  );

  const handleDelete = (id: number) => {
    if (window.confirm('Are you sure you want to delete this entry?')) {
      setData(data.filter((item) => item.id !== id));
    }
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-blue-50">
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
            <button className="bg-[#D4A843] hover:bg-[#b8912f] text-white font-semibold px-6 py-3 rounded-lg shadow-lg transition-all flex items-center gap-2">
              <i className="fa-solid fa-plus"></i> Add New Entry
            </button>
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
            <button className="text-sm text-[#1A3A6B] hover:text-[#D4A843] font-medium flex items-center gap-1 transition-colors">
              <i className="fa-solid fa-filter"></i> Filter
            </button>
          </div>
        </div>

        {/* Table */}
        <div className="bg-white rounded-xl shadow-md overflow-hidden">
          <div className="overflow-x-auto">
            <table className="w-full">
              {/* Table Header with merged sub-header */}
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

              {/* Table Body */}
              <tbody className="divide-y divide-gray-100">
                {filteredData.length === 0 ? (
                  <tr>
                    <td colSpan={6} className="px-4 py-12 text-center text-gray-400">
                      <i className="fa-solid fa-inbox text-4xl mb-3 block"></i>
                      <p className="text-sm">No entries found matching your search.</p>
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
                        <span className="inline-block">
                          {item.remarks}
                        </span>
                      </td>
                      <td className="px-4 py-4 text-center">
                        <div className="flex items-center justify-center gap-2">
                          <button
                            className="w-8 h-8 rounded-lg bg-blue-50 hover:bg-blue-100 text-blue-600 flex items-center justify-center transition-colors"
                            title="View"
                          >
                            <i className="fa-solid fa-eye text-xs"></i>
                          </button>
                          <button
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

          {/* Table Footer / Pagination */}
          <div className="bg-slate-50 px-4 py-3 border-t border-gray-200 flex flex-col md:flex-row items-center justify-between gap-3">
            <p className="text-sm text-gray-500">
              Showing <span className="font-semibold text-[#0B1D3A]">{filteredData.length}</span> of{' '}
              <span className="font-semibold text-[#0B1D3A]">{data.length}</span> entries
            </p>
            <div className="flex items-center gap-1">
              <button className="px-3 py-1.5 text-sm text-gray-500 hover:bg-white rounded border border-transparent hover:border-gray-200 transition-colors">
                <i className="fa-solid fa-chevron-left"></i>
              </button>
              <button className="px-3 py-1.5 text-sm bg-[#0B1D3A] text-white rounded">1</button>
              <button className="px-3 py-1.5 text-sm text-gray-700 hover:bg-white rounded border border-transparent hover:border-gray-200 transition-colors">2</button>
              <button className="px-3 py-1.5 text-sm text-gray-700 hover:bg-white rounded border border-transparent hover:border-gray-200 transition-colors">3</button>
              <button className="px-3 py-1.5 text-sm text-gray-500 hover:bg-white rounded border border-transparent hover:border-gray-200 transition-colors">
                <i className="fa-solid fa-chevron-right"></i>
              </button>
            </div>
          </div>
        </div>

        {/* Legend / Info Box */}
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
