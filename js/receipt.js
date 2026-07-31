// TIEcon Matchmaking Platform
// js/receipt.js — shared "registration record" receipt: fill, print, download as PDF.
// Requires the page to include a #receiptBox with #receiptSub, #receiptRefId,
// #receiptDate and #receiptFields elements (see register-startup.html /
// register-industry.html), plus the jsPDF CDN script for downloadReceiptPDF().

function showReceipt(refId, subtitle, fields) {
  document.getElementById('receiptSub').textContent = subtitle;
  document.getElementById('receiptRefId').textContent = refId;
  document.getElementById('receiptDate').textContent = new Date().toLocaleString('en-IN', { dateStyle: 'medium', timeStyle: 'short' });

  const container = document.getElementById('receiptFields');
  container.innerHTML = '';
  Object.entries(fields).forEach(([label, value]) => {
    if (!value) return;
    const row = document.createElement('div');
    row.className = 'receipt-row';
    const labelSpan = document.createElement('span');
    labelSpan.className = 'rlabel';
    labelSpan.textContent = label;
    const valueSpan = document.createElement('span');
    valueSpan.className = 'rvalue';
    valueSpan.textContent = value;
    row.appendChild(labelSpan);
    row.appendChild(valueSpan);
    container.appendChild(row);
  });

  document.getElementById('receiptBox').style.display = 'block';
  document.getElementById('receiptBox').scrollIntoView({ behavior: 'smooth', block: 'nearest' });
}

function loadImageAsDataURL(src, timeoutMs) {
  const maxWidth = 300; // downscale before embedding - the PDF only needs a small header-sized logo
  const loadPromise = new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      const scale = Math.min(1, maxWidth / img.width);
      const canvas = document.createElement('canvas');
      canvas.width = img.width * scale;
      canvas.height = img.height * scale;
      canvas.getContext('2d').drawImage(img, 0, 0, canvas.width, canvas.height);
      try { resolve(canvas.toDataURL('image/png')); }
      catch (e) { reject(e); }
    };
    img.onerror = reject;
    img.src = src;
  });
  const timeoutPromise = new Promise((_, reject) => setTimeout(() => reject(new Error('logo load timed out')), timeoutMs || 4000));
  return Promise.race([loadPromise, timeoutPromise]);
}

async function downloadReceiptPDF() {
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF();
  const refId = document.getElementById('receiptRefId').textContent;
  const dateStr = document.getElementById('receiptDate').textContent;
  const subtitle = document.getElementById('receiptSub').textContent;

  try {
    const logo = await loadImageAsDataURL('TiE_Vadodara_safe_margin.png');
    doc.addImage(logo, 'PNG', 15, 10, 40, 14);
  } catch (e) { /* logo is cosmetic only - skip if it can't be loaded */ }

  doc.setFontSize(14);
  doc.setFont(undefined, 'bold');
  doc.text('TIEcon Match-a-Thon - Registration Confirmation', 60, 18);
  doc.setFontSize(10);
  doc.setFont(undefined, 'normal');
  doc.text(subtitle, 60, 24);

  let y = 36;
  doc.setDrawColor(230, 57, 70);
  doc.line(15, y, 195, y);
  y += 8;

  doc.setFontSize(10);
  doc.text(`Reference ID: ${refId}`, 15, y);
  doc.text(`Submitted: ${dateStr}`, 120, y);
  y += 10;

  const rows = document.querySelectorAll('#receiptFields .receipt-row');
  rows.forEach(row => {
    const label = row.querySelector('.rlabel').textContent;
    const value = row.querySelector('.rvalue').textContent;
    const valueLines = doc.splitTextToSize(value, 125);

    if (y + valueLines.length * 6 > 285) { doc.addPage(); y = 20; }

    doc.setFont(undefined, 'bold');
    doc.text(`${label}:`, 15, y);
    doc.setFont(undefined, 'normal');
    doc.text(valueLines, 70, y);
    y += Math.max(6, valueLines.length * 6) + 2;
  });

  doc.save(`TIEcon-Registration-${refId}.pdf`);
}
