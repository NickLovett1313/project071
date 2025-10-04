// Client-only Excel comparator using SheetJS (xlsx).
// Strict row-by-row match of 4 fields with cleaning rules.

// Column positions (0-based)
// SAP: C(2)=PO, D(3)=Location, G(6)=Item, H(7)=Model
// OO : N(13)=PO, O(14)=Item, Q(16)=Model, J(9)=Location

const sapCols = { po:2, loc:3, item:6, model:7 };
const ooCols  = { po:13, item:14, model:16, loc:9 };

function cleanModel(s){
  if (s == null) return '';
  const str = String(s);
  const idx = str.indexOf('>>');
  return (idx >= 0 ? str.slice(0, idx) : str).trim();
}

function normPO(s){
  return String(s ?? '').trim().toUpperCase().replaceAll(' ', '');
}

function normLine(s){
  const digits = String(s ?? '').replace(/[^0-9]/g, '');
  const stripped = digits.replace(/^0+/, '');
  return stripped.length ? stripped : '0';
}

function normLocSAP(s){
  const code = String(s ?? '').trim();
  return code === '2913' ? 'EDMONTON' : code.toUpperCase();
}

function normLocOO(s){
  return String(s ?? '').trim().toUpperCase();
}

function readFirstSheet(file){
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try{
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, {type:'array'});
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const rows = XLSX.utils.sheet_to_json(ws, { header:1, blankrows:false });
        resolve(rows);
      }catch(err){ reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

function extractSAP(rows){
  // assume row 0 is header; start from row 1
  const out = [];
  for(let i=1;i<rows.length;i++){
    const r = rows[i] || [];
    const po = normPO(r[sapCols.po]);
    const item = normLine(r[sapCols.item]);
    const model = cleanModel(r[sapCols.model]);
    const loc = normLocSAP(r[sapCols.loc]);
    if(po || item || model || loc){
      out.push({ po, item, model, loc });
    }
  }
  return out;
}

function extractOO(rows){
  const out = [];
  for(let i=1;i<rows.length;i++){
    const r = rows[i] || [];
    const po = normPO(r[ooCols.po]);
    const item = normLine(r[ooCols.item]);
    const model = cleanModel(r[ooCols.model]);
    const loc = normLocOO(r[ooCols.loc]);
    if(po || item || model || loc){
      out.push({ po, item, model, loc });
    }
  }
  return out;
}

function compareRows(sap, oo){
  const n = Math.min(sap.length, oo.length);
  const combined = [];
  const disc = [];
  let eqCount = 0;

  for(let i=0;i<n;i++){
    const s = sap[i], o = oo[i];
    const issues = [];
    if(s.po !== o.po) issues.push(`PO #: SAP=${s.po} vs OO=${o.po}`);
    if(s.item !== o.item) issues.push(`PO Item: SAP=${s.item} vs OO=${o.item}`);
    if(s.model !== o.model) issues.push(`Model: SAP='${s.model}' vs OO='${o.model}'`);
    if(s.loc !== o.loc) issues.push(`Location: SAP='${s.loc}' vs OO='${o.loc}'`);

    combined.push({
      excelRow: i+2, // header row + 1-based
      sap_po: s.po, sap_item: s.item, sap_model: s.model, sap_loc: s.loc,
      oo_po: o.po,  oo_item: o.item,  oo_model: o.model,  oo_loc: o.loc,
      eq: issues.length === 0 ? '✅' : '❌'
    });

    if(issues.length) {
      disc.push({ excelRow: i+2, whatsDifferent: issues.join(' | ') });
    } else {
      eqCount++;
    }
  }
  return { rowsChecked: n, rowsEq: eqCount, disc, combined, sapLen: sap.length, ooLen: oo.length };
}

function renderTable(el, data, columns){
  if(!data || !data.length){ el.innerHTML = '<div class="note">No rows.</div>'; return; }
  const thead = `<thead><tr>${columns.map(c=>`<th>${c.header}</th>`).join('')}</tr></thead>`;
  const tbody = `<tbody>${data.map(row=>{
    return `<tr>${columns.map(c=>`<td>${escapeHtml(row[c.key] ?? '')}</td>`).join('')}</tr>`;
  }).join('')}</tbody>`;
  el.innerHTML = `<table class="table">${thead}${tbody}</table>`;
}

function escapeHtml(s){
  return String(s)
    .replace(/&/g,'&amp;')
    .replace(/</g,'&lt;')
    .replace(/>/g,'&gt;')
    .replace(/\"/g,'&quot;')
    .replace(/'/g,'&#39;');
}

function toCSV(rows, headers){
  const escape = (v)=>`"${String(v ?? '').replace(/"/g,'""')}"`;
  const head = headers.map(h=>escape(h.header)).join(',');
  const lines = rows.map(r=>headers.map(h=>escape(r[h.key])).join(','));
  return [head, ...lines].join('\n');
}

document.getElementById('compareBtn').addEventListener('click', async () => {
  const sapFile = document.getElementById('sapFile').files[0];
  const ooFile = document.getElementById('ooFile').files[0];
  if(!sapFile || !ooFile){
    alert('Please choose both files first.');
    return;
  }

  try{
    const [sapRows, ooRows] = await Promise.all([readFirstSheet(sapFile), readFirstSheet(ooFile)]);
    const sap = extractSAP(sapRows);
    const oo  = extractOO(ooRows);
    const result = compareRows(sap, oo);

    // Summary
    document.getElementById('summary').classList.remove('hidden');
    document.getElementById('rowsChecked').textContent = result.rowsChecked;
    document.getElementById('rowsEq').textContent = result.rowsEq;
    document.getElementById('rowsDisc').textContent = result.rowsChecked - result.rowsEq;

    // Length note if different lengths
    const lengthNote = document.getElementById('lengthNote');
    if(result.sapLen !== result.ooLen){
      lengthNote.classList.remove('hidden');
      lengthNote.textContent = `Files differ in length. SAP rows: ${result.sapLen}, OO rows: ${result.ooLen}. Only the first ${result.rowsChecked} rows were compared (strict row matching).`;
    } else {
      lengthNote.classList.add('hidden');
      lengthNote.textContent = '';
    }

    // Combined table
    renderTable(
      document.getElementById('combinedTable'),
      result.combined,
      [
        { header:'Excel Row', key:'excelRow' },
        { header:'SAP — PO', key:'sap_po' },
        { header:'SAP — Item', key:'sap_item' },
        { header:'SAP — Model', key:'sap_model' },
        { header:'SAP — Location', key:'sap_loc' },
        { header:'OO — PO', key:'oo_po' },
        { header:'OO — Item', key:'oo_item' },
        { header:'OO — Model', key:'oo_model' },
        { header:'OO — Location', key:'oo_loc' },
        { header:'Equivalent?', key:'eq' }
      ]
    );

    // Discrepancy table + CSV
    const discTable = document.getElementById('discTable');
    if(result.disc.length){
      renderTable(discTable, result.disc, [
        { header:'Excel Row', key:'excelRow' },
        { header:`What's Different`, key:'whatsDifferent' }
      ]);
      const csvBtn = document.getElementById('downloadCsv');
      csvBtn.classList.remove('hidden');
      const csv = toCSV(result.disc, [
        { header:'Excel Row', key:'excelRow' },
        { header:`What's Different`, key:'whatsDifferent' }
      ]);
      csvBtn.onclick = () => {
        const blob = new Blob([csv], { type:'text/csv;charset=utf-8;' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'discrepancies.csv';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
      };
    } else {
      discTable.innerHTML = '<div class="note">No discrepancies found.</div>';
      document.getElementById('downloadCsv').classList.add('hidden');
    }
  }catch(err){
    console.error(err);
    alert('Failed to read one of the files. Make sure they are valid Excel files.');
  }
});
