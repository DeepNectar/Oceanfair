import { Link } from 'react-router-dom';

export default function SupplierList() {
  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-blue-50">
      <nav className="bg-[#0B1D3A] shadow-lg">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex items-center justify-between h-16">
            <Link to="/" className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-full bg-white/10 flex items-center justify-center border border-[#D4A843]/30">
                <span className="text-[#D4A843] font-bold text-sm">OF</span>
              </div>
              <span className="text-white font-bold text-sm hidden sm:block">Ocean Fair International</span>
            </Link>
            <Link to="/" className="text-blue-300 hover:text-white text-sm flex items-center gap-2">
              <i className="fa-solid fa-arrow-left"></i> Back to Dashboard
            </Link>
          </div>
        </div>
      </nav>
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-16 text-center">
        <div className="bg-white rounded-2xl shadow-lg p-12 max-w-2xl mx-auto">
          <div className="w-20 h-20 rounded-full bg-rose-100 flex items-center justify-center mx-auto mb-6">
            <i className="fa-solid fa-building text-rose-600 text-3xl"></i>
          </div>
          <h1 className="text-3xl font-bold text-gray-800 mb-4">SUPPLIER LIST</h1>
          <p className="text-gray-500 text-lg">Manage and browse your complete supplier directory here.</p>
          <p className="text-gray-400 text-sm mt-4">This page is under construction. Check back soon!</p>
        </div>
      </div>
    </div>
  );
}
