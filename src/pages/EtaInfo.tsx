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

  // ========== SEND BUTTON — ETA MAILER ==========
  const [isEmailModalOpen, setIsEmailModalOpen] = useState(false);
  const [emailRows, setEmailRows] = useState<
    { sr: number; vessel: string; port: string; eta: string; remarks: string }[]
  >([]);
  const [dispatchMode, setDispatchMode] = useState<'send' | 'display'>('send');

  // Recipients — leave blank for now (user will provide later)
  const RECIPIENTS = {
    to: '',
    cc: '',
  };

  // Port → color class map (mirrors VBA)
  const PORT_COLOR_MAP: Record<string, string> = {
    FUJAIRAH: '#B71C1C',
    DIBBA: '#E65100',
    'KHOR FAKKAN': '#4A148C',
    KFK: '#4A148C',
    'MINA SAQR': '#880E4F',
    'RAS AL KHAIMAH': '#880E4F',
    DUBAI: '#0D47A1',
    'JEBEL ALI': '#1B5E20',
    'JEBAL ALI': '#1B5E20',
    'DUBAI MARITIME CITY': '#004D40',
    DMC: '#004D40',
    SHARJAH: '#F57F17',
    'HAMRIYAH SHARJAH': '#F57F17',
    'ABU DHABI': '#4E342E',
    'ABU DHABI PORT': '#4E342E',
    'KHALIFA PORT': '#4E342E',
    'KHALID PORT': '#F57F17',
    'DRY DOCK': '#00695C',
  };

  const handleSend = () => {
    if (data.length === 0) {
      showNotification(
        'error',
        'No data to send. Please add entries or import from Excel first.'
      );
      return;
    }
    // Populate email rows from current data
    const rows = data.map((d) => ({
      sr: d.srNo,
      vessel: d.vesselName,
      port: d.port,
      eta: d.etaEtbEtd,
      remarks: d.remarks,
    }));
    setEmailRows(rows);
    setIsEmailModalOpen(true);
  };

  const loadSampleEmailRows = () => {
    setEmailRows([
      { sr: 1, vessel: 'MV OCEAN STAR', port: 'FUJAIRAH', eta: '12-Oct 08:00 / 10:00 / 14:00', remarks: 'Confirm for FFV, Bread & Dairy' },
      { sr: 2, vessel: 'MT GULF TRADER', port: 'JEBEL ALI', eta: '12-Oct 14:00 / 16:00 / 13-Oct 06:00', remarks: 'Fresh items required' },
      { sr: 3, vessel: 'MV DESERT PEARL', port: 'ABU DHABI', eta: '13-Oct 06:00 / 08:00 / 18:00', remarks: 'Awaiting agent confirmation' },
      { sr: 4, vessel: 'MT ARABIAN SEA', port: 'SHARJAH', eta: '13-Oct 10:00 / 12:00 / 14-Oct 02:00', remarks: 'Bread & Dairy arrangement' },
      { sr: 5, vessel: 'MV CORAL QUEEN', port: 'KHALIFA PORT', eta: '14-Oct 07:00 / 09:00 / 19:00', remarks: 'Confirm FFV' },
      { sr: 6, vessel: 'MT RAS AL KHAIMAH', port: 'RAS AL KHAIMAH', eta: '14-Oct 15:00 / 17:00 / 15-Oct 05:00', remarks: 'Standard' },
    ]);
    showNotification('success', 'Sample ETA rows loaded.');
  };

  const clearEmailRows = () => {
    setEmailRows([]);
    showNotification('success', 'Email rows cleared.');
  };

  const escapeHtml = (text: string | undefined | null) => {
    if (text === undefined || text === null) return '';
    return String(text)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  };

  const buildEtaTableHtml = () => {
    if (!emailRows.length) {
      return '<p style="text-align:center;color:#94a3b8;font-style:italic;padding:20px 0;font-family:Inter,sans-serif;font-size:13px;">No ETA rows available.</p>';
    }
    let html = '<div style="border-radius:12px;overflow:hidden;box-shadow:0 10px 30px rgba(6,24,43,0.12);border:1px solid rgba(212,175,55,0.3);background:white;margin-bottom:8px;"><table style="border-collapse:collapse;width:100%;font-family:Inter,Arial,sans-serif;"><thead><tr style="background:linear-gradient(135deg,#0b2b4a 0%,#123e63 100%);color:#f0d878;">';
    html += '<th style="padding:15px 12px;text-align:center;font-size:0.72rem;font-weight:700;letter-spacing:0.14em;text-transform:uppercase;border-right:1px solid rgba(212,175,55,0.2);">SR#</th>';
    html += '<th style="padding:15px 12px;text-align:center;font-size:0.72rem;font-weight:700;letter-spacing:0.14em;text-transform:uppercase;border-right:1px solid rgba(212,175,55,0.2);">VESSEL NAME</th>';
    html += '<th style="padding:15px 12px;text-align:center;font-size:0.72rem;font-weight:700;letter-spacing:0.14em;text-transform:uppercase;border-right:1px solid rgba(212,175,55,0.2);">PORT</th>';
    html += '<th style="padding:15px 12px;text-align:center;font-size:0.72rem;font-weight:700;letter-spacing:0.14em;text-transform:uppercase;border-right:1px solid rgba(212,175,55,0.2);">ETA · ETB · ETD</th>';
    html += '<th style="padding:15px 12px;text-align:center;font-size:0.72rem;font-weight:700;letter-spacing:0.14em;text-transform:uppercase;">REMARKS<span style="display:block;font-size:0.62rem;font-weight:500;letter-spacing:0.08em;color:rgba(203,213,225,0.75);text-transform:none;margin-top:4px;line-height:1.4;">Confirm for FFV, Bread &amp; Dairy arrangement</span></th>';
    html += '</tr></thead><tbody>';
    emailRows.forEach((row, idx) => {
      const portUpper = (row.port || '').toString().toUpperCase();
      const portColor = PORT_COLOR_MAP[portUpper] || '#263238';
      const bg = idx % 2 === 1 ? '#fcfbf8' : 'transparent';
      html += `<tr style="background:${bg};">`;
      html += `<td style="padding:14px;text-align:center;font-weight:700;color:#d4af37;font-size:0.82rem;background:${idx % 2 === 1 ? '#f7f4e8' : '#fbfaf5'};width:54px;border-bottom:1px solid #eef2f7;border-right:1px solid #f1f5f9;">${escapeHtml(String(row.sr))}</td>`;
      html += `<td style="padding:14px;font-size:0.84rem;color:#1e2e3f;border-bottom:1px solid #eef2f7;border-right:1px solid #f1f5f9;font-weight:700;color:#0b2b4a;">${escapeHtml(row.vessel)}</td>`;
      html += `<td style="padding:14px;text-align:center;font-weight:700;font-size:0.78rem;letter-spacing:0.04em;color:${portColor};border-bottom:1px solid #eef2f7;border-right:1px solid #f1f5f9;">${escapeHtml(row.port)}</td>`;
      html += `<td style="padding:14px;text-align:center;font-weight:600;color:#0b2b4a;font-size:0.8rem;font-family:Inter,monospace;letter-spacing:0.01em;border-bottom:1px solid #eef2f7;border-right:1px solid #f1f5f9;">${escapeHtml(row.eta)}</td>`;
      html += `<td style="padding:14px;text-align:center;font-size:0.82rem;color:#3e5f7a;font-weight:500;border-bottom:1px solid #eef2f7;">${escapeHtml(row.remarks)}</td>`;
      html += '</tr>';
    });
    html += '</tbody></table></div>';
    return html;
  };

  const buildEmailBodyHtml = () => {
    const tableHtml = buildEtaTableHtml();
    return `
      <div style="background:linear-gradient(135deg,#0b2b4a 0%,#123e63 50%,#0b2b4a 100%);padding:38px 40px 32px;position:relative;overflow:hidden;border-bottom:4px solid #d4af37;text-align:center;">
        <div style="position:absolute;top:-50%;right:-10%;width:400px;height:400px;background:radial-gradient(circle,rgba(212,175,55,0.18),transparent 65%);border-radius:50%;pointer-events:none;"></div>
        <h1 style="font-family:'Playfair Display',serif;font-size:32pt;font-weight:800;color:#ffffff;letter-spacing:-0.015em;line-height:1.05;margin:0 0 10px;position:relative;z-index:1;">Vessel <span style="background:linear-gradient(120deg,#f0d878 0%,#d4af37 100%);-webkit-background-clip:text;background-clip:text;color:transparent;">Schedule</span> Update</h1>
        <div style="font-size:11px;letter-spacing:0.32em;color:rgba(240,216,120,0.85);text-transform:uppercase;font-weight:600;position:relative;z-index:1;">ETA · ETB · ETD Notification</div>
      </div>
      <div style="padding:34px 40px 36px;background:#fbfaf7;font-family:'Cormorant Garamond','Times New Roman',serif;">
        <div style="font-family:'Playfair Display',serif;font-size:20px;font-weight:600;color:#0b2b4a;margin-bottom:12px;letter-spacing:-0.01em;">Dear Team,</div>
        <p style="font-family:Inter,sans-serif;font-size:14px;line-height:1.7;color:#3e5f7a;margin-bottom:26px;font-weight:400;">
          Please advise the exact <b style="color:#0b2b4a;font-weight:600;background:linear-gradient(180deg,transparent 60%,rgba(212,175,55,0.25) 60%);padding:0 2px;">ETA</b> for the below vessels with <b style="color:#0b2b4a;font-weight:600;background:linear-gradient(180deg,transparent 60%,rgba(212,175,55,0.25) 60%);padding:0 2px;">REMARKS</b> to enable us to arrange the fresh items accordingly.
        </p>
        <div style="display:flex;align-items:center;gap:14px;margin:24px 0 22px;">
          <div style="flex:1;height:1px;background:linear-gradient(90deg,transparent,rgba(212,175,55,0.5),transparent);"></div>
          <span style="color:#d4af37;font-size:10px;letter-spacing:0.4em;">◆ ◆ ◆</span>
          <div style="flex:1;height:1px;background:linear-gradient(90deg,transparent,rgba(212,175,55,0.5),transparent);"></div>
        </div>
        ${tableHtml}
      </div>
      <div style="height:6px;background:linear-gradient(90deg,#0b2b4a 0%,#d4af37 50%,#0b2b4a 100%);"></div>
    `;
  };

  const getSubjectLine = () => {
    const today = new Date();
    const subjectDate = today.toLocaleDateString('en-US', {
      weekday: 'long',
      year: 'numeric',
      month: 'long',
      day: 'numeric',
    });
    return `ETA/ETB - INFO -- ${subjectDate}`;
  };

  const buildPlainTextBody = () => {
    let text = 'Dear Team,\n\n';
    text += 'Please advise the exact ETA for the below vessels with REMARKS to enable us to arrange the fresh items accordingly.\n\n';
    text += '═══════════════════════════════════════════════════════\n\n';
    
    // Table header
    text += 'SR#\tVESSEL NAME\tPORT\tETA · ETB · ETD\tREMARKS\n';
    text += '─────────────────────────────────────────────────────────────────────────────────────────────────────────\n';
    
    // Table rows
    emailRows.forEach((row) => {
      text += `${row.sr}\t${row.vessel}\t${row.port}\t${row.eta}\t${row.remarks}\n`;
    });
    
    text += '\n═══════════════════════════════════════════════════════\n\n';
    text += 'Best regards,\n';
    text += 'Oceanfair Operations Team\n';
    text += 'Vessel Coordination\n';
    text += 'Dubai · Fujairah · Abu Dhabi — United Arab Emirates\n';
    
    return text;
  };

  const buildFullHtmlEmail = () => {
    const bodyHtml = buildEmailBodyHtml();
    return `<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Playfair+Display:wght@600;700;800&family=Cormorant+Garamond:wght@500;600;700&display=swap" rel="stylesheet">
<style>
  *{box-sizing:border-box;margin:0;padding:0;}
  body{background:#ffffff;padding:0;font-family:'Inter',Arial,sans-serif;}
</style>
</head>
<body>${bodyHtml}</body>
</html>`;
  };

  const buildMailtoLink = (subject: string) => {
    // Properly construct mailto URL to avoid subject leaking into CC
    const toPart = RECIPIENTS.to || '';
    const params = new URLSearchParams();
    if (RECIPIENTS.cc) params.set('cc', RECIPIENTS.cc);
    params.set('subject', subject);
    return `mailto:${toPart}?${params.toString()}`;
  };

  const dispatchEmail = () => {
    if (!emailRows.length) {
      showNotification('error', 'No ETA rows to send. Click "Load Sample Rows" or add entries first.');
      return;
    }
    const subject = getSubjectLine();
    
    if (dispatchMode === 'display') {
      // Open & Display mode - opens HTML preview window
      openEmailPreviewWindow(subject);
      showNotification('success', 'Email opened in Display mode — review & send manually.');
    } else {
      // Send Directly mode - copies HTML to clipboard + opens default email app
      const htmlEmail = buildFullHtmlEmail();
      const plainTextBody = buildPlainTextBody();
      const mailtoLink = buildMailtoLink(subject);
      
      // Close the modal first
      setIsEmailModalOpen(false);
      
      // Wait for modal to close, then copy HTML to clipboard and open email app
      setTimeout(() => {
        // Try to copy HTML to clipboard (works in modern browsers)
        if (navigator.clipboard && window.ClipboardItem) {
          const htmlBlob = new Blob([htmlEmail], { type: 'text/html' });
          const textBlob = new Blob([plainTextBody], { type: 'text/plain' });
          const clipboardItem = new ClipboardItem({
            'text/html': htmlBlob,
            'text/plain': textBlob,
          });
          
          navigator.clipboard.write([clipboardItem]).then(() => {
            // HTML copied successfully - now open email app
            const link = document.createElement('a');
            link.href = mailtoLink;
            link.style.display = 'none';
            document.body.appendChild(link);
            link.click();
            setTimeout(() => document.body.removeChild(link), 100);
            
            showNotification('success', '✓ HTML email copied! Paste (Ctrl+V / Cmd+V) in your email body.');
          }).catch(() => {
            // Clipboard failed - fallback to plain text mailto
            fallbackPlainMailto(subject, plainTextBody, mailtoLink);
          });
        } else {
          // Clipboard API not available - fallback to plain text
          fallbackPlainMailto(subject, plainTextBody, mailtoLink);
        }
      }, 300);
    }
  };

  const fallbackPlainMailto = (subject: string, plainTextBody: string, mailtoLink: string) => {
    const params = new URLSearchParams();
    if (RECIPIENTS.cc) params.set('cc', RECIPIENTS.cc);
    params.set('subject', subject);
    params.set('body', plainTextBody);
    const toPart = RECIPIENTS.to || '';
    const fullMailto = `mailto:${toPart}?${params.toString()}`;
    
    const link = document.createElement('a');
    link.href = fullMailto;
    link.style.display = 'none';
    document.body.appendChild(link);
    link.click();
    setTimeout(() => document.body.removeChild(link), 100);
    
    showNotification('success', 'Opening your default email app with plain text body...');
  };

  const openEmailPreviewWindow = (subject: string) => {
    const bodyHtml = buildEmailBodyHtml();
    const fullDoc = `<!DOCTYPE html>
<html><head><meta charset="UTF-8"><title>${escapeHtml(subject)}</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Playfair+Display:wght@600;700;800&family=Cormorant+Garamond:wght@500;600;700&display=swap" rel="stylesheet">
<style>
  *{box-sizing:border-box;margin:0;padding:0;}
  body{background:#eef2f7;padding:30px 15px;font-family:'Inter',Arial,sans-serif;}
</style></head><body>
  <div style="max-width:860px;margin:0 auto;background:#ffffff;border-radius:16px;overflow:hidden;box-shadow:0 20px 60px rgba(6,24,43,0.18);font-family:'Inter',Arial,sans-serif;border:1px solid rgba(212,175,55,0.25);">${bodyHtml}</div>
</body></html>`;
    const w = window.open('', '_blank', 'width=960,height=800,scrollbars=yes,resizable=yes');
    if (w) {
      w.document.open();
      w.document.write(fullDoc);
      w.document.close();
      w.document.title = subject;
    } else {
      showNotification('error', 'Popup blocked. Please allow popups to preview the email.');
    }
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

      {/* EMAIL PREVIEW MODAL */}
      {isEmailModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/70 backdrop-blur-sm">
          <div className="bg-gradient-to-br from-[#050f1d] to-[#06182b] rounded-2xl shadow-2xl w-full max-w-6xl max-h-[95vh] overflow-hidden flex flex-col border border-[#D4A843]/20">
            {/* Modal Header */}
            <div className="bg-gradient-to-r from-[#0B1D3A] to-[#1A3A6B] px-6 py-4 flex items-center justify-between flex-shrink-0 border-b border-[#D4A843]/20">
              <div className="flex items-center gap-3">
                <div className="w-12 h-12 rounded-xl bg-gradient-to-br from-[#D4A843] to-[#b8942a] flex items-center justify-center text-2xl">
                  ✉
                </div>
                <div>
                  <h3 className="text-white font-bold text-xl">ETA Mailer</h3>
                  <p className="text-blue-200 text-xs">
                    Choose your dispatch method — send directly or open to review manually.
                  </p>
                </div>
              </div>
              <button
                onClick={() => setIsEmailModalOpen(false)}
                className="text-white/70 hover:text-white transition-colors text-2xl"
              >
                <i className="fa-solid fa-xmark"></i>
              </button>
            </div>

            {/* Mode Selector Chips */}
            <div className="bg-[#0a1a2f]/50 px-6 py-3 flex flex-wrap items-center gap-3 border-b border-[#D4A843]/10">
              <div className="inline-flex items-center gap-1 p-1 rounded-full bg-[#06182b]/70 border border-[#D4A843]/20">
                <button
                  onClick={() => setDispatchMode('send')}
                  className={`inline-flex items-center gap-2 px-4 py-2 rounded-full text-xs font-semibold transition-all ${
                    dispatchMode === 'send'
                      ? 'bg-gradient-to-r from-[#D4A843] to-[#b8942a] text-[#0B1D3A] shadow-lg'
                      : 'text-gray-400 hover:text-white hover:bg-white/10'
                  }`}
                >
                  <span className={`w-2 h-2 rounded-full ${dispatchMode === 'send' ? 'bg-[#0B1D3A]' : 'bg-current'}`}></span>
                  Send Directly
                </button>
                <button
                  onClick={() => setDispatchMode('display')}
                  className={`inline-flex items-center gap-2 px-4 py-2 rounded-full text-xs font-semibold transition-all ${
                    dispatchMode === 'display'
                      ? 'bg-gradient-to-r from-[#D4A843] to-[#b8942a] text-[#0B1D3A] shadow-lg'
                      : 'text-gray-400 hover:text-white hover:bg-white/10'
                  }`}
                >
                  <span className={`w-2 h-2 rounded-full ${dispatchMode === 'display' ? 'bg-[#0B1D3A]' : 'bg-current'}`}></span>
                  Open &amp; Display
                </button>
              </div>

              {/* Toolbar Buttons */}
              <button
                onClick={dispatchEmail}
                className="bg-gradient-to-r from-[#D4A843] to-[#b8942a] hover:from-[#e5c452] hover:to-[#D4A843] text-[#0B1D3A] font-bold px-5 py-2.5 rounded-lg shadow-lg transition-all flex items-center gap-2 ml-auto"
              >
                <span>✉</span> {dispatchMode === 'send' ? 'Send Email' : 'Open & Display'}
              </button>
              <button
                onClick={loadSampleEmailRows}
                className="bg-white/5 hover:bg-white/10 text-white border border-white/20 hover:border-[#D4A843]/40 px-4 py-2.5 rounded-lg transition-all flex items-center gap-2 text-sm font-semibold"
              >
                <span>◈</span> Load Sample
              </button>
              <button
                onClick={clearEmailRows}
                className="bg-transparent hover:bg-red-500/10 text-gray-400 hover:text-red-400 border border-gray-400/20 hover:border-red-400/40 px-4 py-2.5 rounded-lg transition-all flex items-center gap-2 text-sm font-semibold"
              >
                <span>✕</span> Clear
              </button>
            </div>

            {/* Email Preview Content */}
            <div className="flex-1 overflow-y-auto p-6 bg-[#eef2f7]">
              <div className="max-w-4xl mx-auto bg-white rounded-xl shadow-2xl overflow-hidden border border-[#D4A843]/25">
                {/* Email Meta */}
                <div className="bg-[#06182b]/70 px-6 py-4 border-b border-[#D4A843]/10 text-sm space-y-2">
                  <div className="flex gap-3">
                    <span className="font-bold text-[#D4A843] text-xs uppercase tracking-wider min-w-[60px]">To</span>
                    <span className="text-gray-300 font-medium">
                      {RECIPIENTS.to || <span className="text-gray-500 italic">— (To be configured)</span>}
                    </span>
                  </div>
                  <div className="flex gap-3">
                    <span className="font-bold text-[#D4A843] text-xs uppercase tracking-wider min-w-[60px]">CC</span>
                    <span className="text-gray-300 font-medium">
                      {RECIPIENTS.cc || <span className="text-gray-500 italic">— (To be configured)</span>}
                    </span>
                  </div>
                  <div className="flex gap-3">
                    <span className="font-bold text-[#D4A843] text-xs uppercase tracking-wider min-w-[60px]">Subject</span>
                    <span className="text-[#f0d878] font-semibold">{getSubjectLine()}</span>
                  </div>
                </div>

                {/* Email Body */}
                <div className="max-h-[600px] overflow-y-auto">
                  <div
                    className="email-preview-content"
                    dangerouslySetInnerHTML={{ __html: buildEmailBodyHtml() }}
                  />
                </div>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
