import { BrowserRouter as Router, Routes, Route } from 'react-router-dom';
import HomePage from './pages/HomePage';
import EtaInfo from './pages/EtaInfo';
import ReOrder from './pages/ReOrder';
import SendInquiry from './pages/SendInquiry';
import EmailLog from './pages/EmailLog';
import SupplierList from './pages/SupplierList';

export default function App() {
  return (
    <Router>
      <Routes>
        <Route path="/" element={<HomePage />} />
        <Route path="/eta-info" element={<EtaInfo />} />
        <Route path="/re-order" element={<ReOrder />} />
        <Route path="/send-inquiry" element={<SendInquiry />} />
        <Route path="/email-log" element={<EmailLog />} />
        <Route path="/supplier-list" element={<SupplierList />} />
      </Routes>
    </Router>
  );
}
