import { Link } from 'react-router-dom';

const navItems = [
  { path: '/eta-info', label: 'ETA INFO', icon: 'fa-solid fa-ship', description: 'Track estimated arrival times for all shipments', color: 'from-blue-600 to-blue-800' },
  { path: '/re-order', label: 'RE-ORDER', icon: 'fa-solid fa-rotate', description: 'Quick re-order from previous purchase history', color: 'from-emerald-600 to-emerald-800' },
  { path: '/send-inquiry', label: 'SEND INQ TO SUPPLIER', icon: 'fa-solid fa-paper-plane', description: 'Send inquiries directly to your suppliers', color: 'from-amber-600 to-amber-800' },
  { path: '/email-log', label: 'EMAIL LOG', icon: 'fa-solid fa-envelope-open-text', description: 'View all email correspondence and history', color: 'from-purple-600 to-purple-800' },
  { path: '/supplier-list', label: 'SUPPLIER LIST', icon: 'fa-solid fa-building', description: 'Manage and browse your supplier directory', color: 'from-rose-600 to-rose-800' },
];

export default function HomePage() {
  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-blue-50">
      {/* Navigation Bar */}
      <nav className="bg-[#0B1D3A] shadow-lg">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex items-center justify-between h-20">
            {/* Logo & Company Name */}
            <div className="flex items-center gap-4">
              <div className="w-16 h-16 rounded-full bg-white flex items-center justify-center overflow-hidden border-2 border-[#D4A843]/50 shadow-lg">
                <img
                  src="https://image.qwenlm.ai/generated-images/7ae76d05-28af-4a87-952f-4b24a40a2cd7/_result.png"
                  alt="Ocean Fair International Group FZE Logo"
                  className="w-14 h-14 object-contain"
                  onError={(e) => {
                    const target = e.target as HTMLImageElement;
                    target.onerror = null;
                    target.src = "https://lh3.googleusercontent.com/d/1WBc2Su-36a-le49VAP9rS4pRCOjO1qxw";
                  }}
                />
              </div>
              <div>
                <h1 className="text-white text-xl font-bold tracking-wide">Ocean Fair International Group FZE</h1>
                <p className="text-blue-300 text-xs tracking-widest uppercase">Procurement Management Portal</p>
              </div>
            </div>
            {/* Status Indicator */}
            <div className="hidden md:flex items-center gap-2">
              <span className="w-2.5 h-2.5 bg-green-400 rounded-full animate-pulse"></span>
              <span className="text-green-300 text-sm font-medium">System Online</span>
            </div>
          </div>
        </div>
      </nav>

      {/* Hero Section */}
      <section className="relative overflow-hidden">
        <div className="absolute inset-0 bg-gradient-to-r from-[#0B1D3A] via-[#1A3A6B] to-[#0B1D3A]"></div>
        <div className="absolute inset-0 opacity-10">
          <div className="absolute inset-0" style={{
            backgroundImage: `url("data:image/svg+xml,%3Csvg width='60' height='60' viewBox='0 0 60 60' xmlns='http://www.w3.org/2000/svg'%3E%3Cg fill='none' fill-rule='evenodd'%3E%3Cg fill='%23ffffff' fill-opacity='0.4'%3E%3Cpath d='M36 34v-4h-2v4h-4v2h4v4h2v-4h4v-2h-4zm0-30V0h-2v4h-4v2h4v4h2V6h4V4h-4zM6 34v-4H4v4H0v2h4v4h2v-4h4v-2H6zM6 4V0H4v4H0v2h4v4h2V6h4V4H6z'/%3E%3C/g%3E%3C/g%3E%3C/svg%3E")`,
          }}></div>
        </div>
        <div className="relative max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-16 text-center">
          {/* Large Logo */}
          <div className="flex justify-center mb-8">
            <div className="w-28 h-28 md:w-36 md:h-36 rounded-full bg-white/10 backdrop-blur-sm flex items-center justify-center border-2 border-[#D4A843]/40 shadow-2xl">
              <img
                src="https://image.qwenlm.ai/generated-images/7ae76d05-28af-4a87-952f-4b24a40a2cd7/_result.png"
                alt="Ocean Fair International Group FZE"
                className="w-24 h-24 md:w-32 md:h-32 object-contain"
              />
            </div>
          </div>
          <div className="inline-flex items-center gap-2 bg-white/10 backdrop-blur-sm rounded-full px-4 py-2 mb-6">
            <span className="w-2 h-2 bg-[#D4A843] rounded-full"></span>
            <span className="text-blue-100 text-sm font-medium">Welcome to your procurement dashboard</span>
          </div>
          <h2 className="text-4xl md:text-5xl font-bold text-white mb-4">
            Manage Your <span className="text-[#D4A843]">Global Supply Chain</span>
          </h2>
          <p className="text-blue-200 text-lg max-w-2xl mx-auto">
            Streamline your procurement operations — track shipments, manage suppliers, and handle orders all from one centralized platform.
          </p>
        </div>
      </section>

      {/* Dashboard Cards */}
      <section className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 -mt-8 pb-16 relative z-10">
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
          {navItems.map((item) => (
            <Link
              key={item.path}
              to={item.path}
              className="group relative bg-white rounded-2xl shadow-lg hover:shadow-2xl transition-all duration-300 overflow-hidden transform hover:-translate-y-1"
            >
              {/* Card Top Gradient Bar */}
              <div className={`h-1.5 bg-gradient-to-r ${item.color}`}></div>
              
              <div className="p-6">
                {/* Icon */}
                <div className={`w-14 h-14 rounded-xl bg-gradient-to-br ${item.color} flex items-center justify-center mb-4 group-hover:scale-110 transition-transform duration-300`}>
                  <i className={`${item.icon} text-white text-xl`}></i>
                </div>
                
                {/* Content */}
                <h3 className="text-lg font-bold text-gray-800 mb-2 group-hover:text-[#1A3A6B] transition-colors">
                  {item.label}
                </h3>
                <p className="text-gray-500 text-sm leading-relaxed">
                  {item.description}
                </p>
                
                {/* Arrow */}
                <div className="mt-4 flex items-center text-gray-400 group-hover:text-[#1A3A6B] transition-colors">
                  <span className="text-sm font-medium mr-2">Access Module</span>
                  <i className="fa-solid fa-arrow-right text-xs group-hover:translate-x-1 transition-transform"></i>
                </div>
              </div>
            </Link>
          ))}
        </div>
      </section>

      {/* Stats Section */}
      <section className="bg-white border-t border-gray-100">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-12">
          <div className="grid grid-cols-2 md:grid-cols-4 gap-8 text-center">
            <div>
              <div className="text-3xl font-bold text-[#0B1D3A]">24/7</div>
              <div className="text-sm text-gray-500 mt-1">System Availability</div>
            </div>
            <div>
              <div className="text-3xl font-bold text-[#0B1D3A]">100+</div>
              <div className="text-sm text-gray-500 mt-1">Active Suppliers</div>
            </div>
            <div>
              <div className="text-3xl font-bold text-[#0B1D3A]">Real-time</div>
              <div className="text-sm text-gray-500 mt-1">Shipment Tracking</div>
            </div>
            <div>
              <div className="text-3xl font-bold text-[#0B1D3A]">Global</div>
              <div className="text-sm text-gray-500 mt-1">Trade Network</div>
            </div>
          </div>
        </div>
      </section>

      {/* Footer */}
      <footer className="bg-[#0B1D3A] text-white">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
          <div className="flex flex-col md:flex-row items-center justify-between gap-4">
            <div className="flex items-center gap-3">
              <div className="w-10 h-10 rounded-full bg-white/10 flex items-center justify-center overflow-hidden border border-[#D4A843]/30">
                <img
                  src="https://image.qwenlm.ai/generated-images/7ae76d05-28af-4a87-952f-4b24a40a2cd7/_result.png"
                  alt="Ocean Fair Logo"
                  className="w-9 h-9 object-contain"
                />
              </div>
              <div>
                <p className="font-semibold text-sm">Ocean Fair International Group FZE</p>
                <p className="text-blue-300 text-xs">Procurement Management Portal</p>
              </div>
            </div>
            <div className="text-center md:text-right">
              <p className="text-blue-300 text-sm">© {new Date().getFullYear()} Ocean Fair International Group FZE. All rights reserved.</p>
            </div>
          </div>
        </div>
      </footer>
    </div>
  );
}
