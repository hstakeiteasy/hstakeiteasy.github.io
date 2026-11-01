/*******************************************************
 * TUITION MANAGEMENT – APPS SCRIPT BACKEND (API)
 *******************************************************/

const SHEET_NAMES = {
  STUDENTS: 'Students', 
  BATCHES: 'Batches', 
  BILLS: 'Bills',
  PAYMENTS: 'Payments', 
  DISCOUNTS: 'Discounts', 
  SETTINGS: 'Settings', 
  LOGS: 'Logs'
};

const HEADERS = {
  STUDENTS: ['StudentID', 'Name', 'ParentName', 'ParentPhone', 'BatchID', 'CustomFee', 'Password', 'Status', 'JoinDate', 'CreatedAt'],
  BATCHES: ['BatchID', 'BatchName', 'DefaultFee', 'DueDay', 'Active', 'CreatedAt'],
  BILLS: ['BillID', 'StudentID', 'BatchID', 'Month', 'Year', 'Amount', 'DueDate', 'Discount', 'Paid', 'Status', 'Note', 'CreatedAt'],
  PAYMENTS: ['PaymentID', 'StudentID', 'BillID', 'Amount', 'PaidOn', 'Mode', 'Note', 'CreatedAt'],
  DISCOUNTS: ['DiscountID', 'StudentID', 'BillID', 'Amount', 'Note', 'CreatedAt'],
  SETTINGS: ['Key', 'Value'],
  LOGS: ['Timestamp', 'User', 'Action', 'Details']
};

/* ---------------- INIT: Create sheets & headers ---------------- */
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  Object.keys(SHEET_NAMES).forEach(key => {
    let sh = ss.getSheetByName(SHEET_NAMES[key]);
    if (!sh) {
      sh = ss.insertSheet(SHEET_NAMES[key]);
    }
    
    const headers = HEADERS[key];
    const existing = sh.getRange(1, 1, 1, headers.length).getValues()[0];
    
    if (existing.every(v => v === '')) {
      sh.getRange(1, 1, 1, headers.length)
        .setValues([headers])
        .setFontWeight('bold')
        .setBackground('#f1f5f9');
      sh.setFrozenRows(1);
    }
  });
  
  const setSh = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (setSh.getLastRow() < 2) {
    setSh.appendRow(['admin_user', 'admin']); 
    setSh.appendRow(['admin_pass', 'admin123']);
  }
}

/* ---------------- ROUTER ---------------- */
function doGet(e) {
  initSheets();
  const action = e.parameter.action;
  
  try {
    let result;
    switch (action) {
      case 'login':                      result = apiLogin(e.parameter); break;
      case 'me':                         result = apiMe(e.parameter); break;
      case 'getBatches':                 result = apiGetBatches(e.parameter); break;
      case 'getStudents':                result = apiGetStudents(e.parameter); break;
      case 'addStudent':                 result = apiAddStudent(e.parameter); break;
      case 'updateStudent':              result = apiUpdateStudent(e.parameter); break;
      case 'adminChangeStudentPassword': result = apiAdminChangeStudentPassword(e.parameter); break;
      case 'addBatch':                   result = apiAddBatch(e.parameter); break;
      case 'updateBatch':                result = apiUpdateBatch(e.parameter); break;
      case 'studentFees':                result = apiStudentFees(e.parameter); break;
      case 'batchDues':                  result = apiBatchDues(e.parameter); break;
      case 'collectMultiFee':            result = apiCollectMultiFee(e.parameter); break;
      case 'applyMultiDiscount':         result = apiApplyMultiDiscount(e.parameter); break;
      case 'generateBills':              result = apiGenerateBills(e.parameter); break;
      case 'deleteBills':                result = apiDeleteBills(e.parameter); break;
      case 'collectionReport':           result = apiCollectionReport(e.parameter); break;
      case 'addPrevDue':                 result = apiAddPrevDue(e.parameter); break;
      default: result = {ok: false, msg: 'Unknown action'};
    }
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ok: false, msg: String(err)})).setMimeType(ContentService.MimeType.JSON);
  }
}

/* ---------------- HELPERS ---------------- */
function ss_() { 
  return SpreadsheetApp.getActiveSpreadsheet(); 
}

function sheet_(name) { 
  return ss_().getSheetByName(name); 
}

function readRows_(sheetName) {
  const sh = sheet_(sheetName); 
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return [];
  
  const head = v[0]; 
  return v.slice(1).map(r => { 
    const o = {}; 
    head.forEach((h, i) => o[h] = r[i]); 
    return o; 
  });
}

function append_(sheetName, obj) {
  const sh = sheet_(sheetName); 
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const row = head.map(h => obj[h] !== undefined ? obj[h] : '');
  sh.appendRow(row);
}

function updateRow_(sheetName, keyField, keyValue, updates) {
  const sh = sheet_(sheetName); 
  const v = sh.getDataRange().getValues();
  const head = v[0]; 
  const keyIdx = head.indexOf(keyField);
  
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][keyIdx]) === String(keyValue)) {
      Object.keys(updates).forEach(k => {
        const ci = head.indexOf(k); 
        if (ci > -1) {
          sh.getRange(i + 1, ci + 1).setValue(updates[k]);
        }
      });
      return true;
    }
  } 
  return false;
}

function now_() { 
  return new Date(); 
}

function istString_(d) { 
  return Utilities.formatDate(d, 'Asia/Kolkata', 'yyyy-MM-dd HH:mm:ss'); 
}

function randomLetters_(n) { 
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ'; 
  let s = ''; 
  for(let i = 0; i < n; i++) {
    s += chars[Math.floor(Math.random() * chars.length)]; 
  }
  return s; 
}

function getSetting_(k) { 
  const r = readRows_(SHEET_NAMES.SETTINGS).find(x => x.Key === k); 
  return r ? r.Value : ''; 
}

/* ---------------- AUTH ---------------- */
function apiLogin(p) {
  if (p.role === 'admin') {
    if (p.username === getSetting_('admin_user') && p.password === getSetting_('admin_pass')) {
      const t = Utilities.getUuid(); 
      CacheService.getScriptCache().put('admin_' + t, '1', 21600);
      return {ok: true, role: 'admin', token: t, name: 'Admin'};
    } 
    return {ok: false, msg: 'Invalid admin credentials'};
  }
  
  if (p.role === 'student') {
    const st = readRows_(SHEET_NAMES.STUDENTS).find(s => 
      String(s.StudentID).toUpperCase() === String(p.username).toUpperCase() && 
      String(s.Password) === String(p.password)
    );
    
    if (!st) return {ok: false, msg: 'Invalid student ID or password'};
    
    const t = Utilities.getUuid(); 
    CacheService.getScriptCache().put('stu_' + t, st.StudentID, 21600);
    return {ok: true, role: 'student', token: t, studentId: st.StudentID, name: st.Name};
  } 
  
  return {ok: false, msg: 'Unknown role'};
}

function apiMe(p) { 
  if (p.role === 'admin' && CacheService.getScriptCache().get('admin_' + p.token)) {
    return {ok: true, role: 'admin'}; 
  }
  
  if (p.role === 'student') {
    const sid = CacheService.getScriptCache().get('stu_' + p.token);
    if (sid) {
      const st = readRows_(SHEET_NAMES.STUDENTS).find(s => s.StudentID === sid); 
      if (st) {
        return {ok: true, role: 'student', studentId: st.StudentID, name: st.Name, batchId: st.BatchID}; 
      }
    }
  } 
  return {ok: false}; 
}

/* ---------------- BATCHES ---------------- */
function apiGetBatches() { 
  return {ok: true, data: readRows_(SHEET_NAMES.BATCHES)}; 
}

function apiAddBatch(p) {
  const rows = readRows_(SHEET_NAMES.BATCHES);
  const id = 'B' + (rows.length + 1);
  
  append_(SHEET_NAMES.BATCHES, { 
    BatchID: id, 
    BatchName: p.name, 
    DefaultFee: Number(p.fee), 
    DueDay: Number(p.dueDay) || 5, 
    Active: 'YES', 
    CreatedAt: istString_(now_()) 
  });
  
  return {ok: true};
}

function apiUpdateBatch(p) {
  updateRow_(SHEET_NAMES.BATCHES, 'BatchID', p.batchId, { 
    BatchName: p.name, 
    DefaultFee: Number(p.fee), 
    DueDay: Number(p.dueDay) 
  });
  return {ok: true}; 
}

/* ---------------- STUDENTS ---------------- */
function apiGetStudents(p) { 
  return {ok: true, data: readRows_(SHEET_NAMES.STUDENTS)}; 
}

function apiAddStudent(p) {
  const name = String(p.name).trim(); 
  const rows = readRows_(SHEET_NAMES.STUDENTS); 
  const sid = 'TIE' + (rows.length + 1);
  
  const phone = String(p.parentPhone || '').replace(/\D/g, '');
  const pass = name.slice(0, 4).padEnd(4, 'X') + '@' + phone.slice(-4).padStart(4, '0') + randomLetters_(2);
  
  append_(SHEET_NAMES.STUDENTS, { 
    StudentID: sid, 
    Name: name, 
    ParentName: p.parentName || '', 
    ParentPhone: phone, 
    BatchID: p.batchId || '', 
    CustomFee: p.customFee ? Number(p.customFee) : '', 
    Password: pass, 
    Status: 'ACTIVE', 
    JoinDate: istString_(now_()).slice(0, 10), 
    CreatedAt: istString_(now_()) 
  });
  
  return {ok: true, studentId: sid, password: pass, parentPhone: phone};
}

function apiUpdateStudent(p) {
  updateRow_(SHEET_NAMES.STUDENTS, 'StudentID', p.studentId, { 
    Name: p.name, 
    ParentPhone: p.parentPhone, 
    CustomFee: p.customFee ? Number(p.customFee) : '', 
    BatchID: p.batchId 
  });
  return {ok: true}; 
}

function apiAdminChangeStudentPassword(p) {
  updateRow_(SHEET_NAMES.STUDENTS, 'StudentID', p.studentId, { 
    Password: p.newPassword 
  });
  return {ok: true}; 
}

/* ---------------- BILLS & FEES ---------------- */
function apiGenerateBills(p) {
  // VERIFY PASSWORD
  if (p.adminPass !== getSetting_('admin_pass')) {
    return {ok: false, msg: 'Incorrect Admin Password'};
  }

  const batches = readRows_(SHEET_NAMES.BATCHES); 
  const students = readRows_(SHEET_NAMES.STUDENTS);
  
  const m = Number(p.month); 
  const y = Number(p.year);
  
  let targets = students.filter(s => s.Status === 'ACTIVE');
  if (p.batchId && p.batchId.trim() !== '') {
    targets = targets.filter(s => s.BatchID === p.batchId);
  }

  targets.forEach(st => {
    // DUPLICATE CHECK REMOVED - You can now generate multiple bills for a single month
    
    const batch = batches.find(b => b.BatchID === st.BatchID);
    
    // Accurately pulls custom fee (Even if custom fee is set to 0)
    let amt;
    if (st.CustomFee !== '' && st.CustomFee !== undefined) {
      amt = Number(st.CustomFee);
    } else {
      amt = batch ? Number(batch.DefaultFee) : 0;
    }
    
    const due = new Date(y, m - 1, batch ? Number(batch.DueDay) : 5);
    const bid = 'BILL-' + Utilities.formatDate(now_(), 'Asia/Kolkata', 'yyMMddHHmmss') + '-' + st.StudentID + '-' + randomLetters_(2);
    
    append_(SHEET_NAMES.BILLS, { 
      BillID: bid, 
      StudentID: st.StudentID, 
      BatchID: st.BatchID, 
      Month: m, 
      Year: y, 
      Amount: amt, 
      DueDate: Utilities.formatDate(due, 'Asia/Kolkata', 'yyyy-MM-dd'), 
      Discount: 0, 
      Paid: 0, 
      Status: 'UNPAID', 
      Note: '', 
      CreatedAt: istString_(now_()) 
    });
  }); 
  
  return {ok: true};
}

function apiDeleteBills(p) {
  // VERIFY PASSWORD
  if (p.adminPass !== getSetting_('admin_pass')) {
    return {ok: false, msg: 'Incorrect Admin Password'};
  }

  const sh = sheet_(SHEET_NAMES.BILLS);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return {ok: true, count: 0};
  
  const head = data[0];
  const iBatch = head.indexOf('BatchID'); 
  const iMonth = head.indexOf('Month'); 
  const iYear = head.indexOf('Year');
  const iStudent = head.indexOf('StudentID');
  
  let deleted = 0;
  let deletedStudents = new Set();
  
  // Loop backwards to cleanly delete "Stack by Stack" (one bill per student per click)
  for (let i = data.length - 1; i >= 1; i--) {
    const row = data[i];
    const matchBatch = (p.batchId && p.batchId.trim() !== '') ? row[iBatch] === p.batchId : true; 
    
    if (matchBatch && String(row[iMonth]) === String(p.month) && String(row[iYear]) === String(p.year)) {
      const sid = row[iStudent];
      
      // If we haven't deleted a bill for this student in this loop, delete it now
      if (!deletedStudents.has(sid)) {
        sh.deleteRow(i + 1);
        deletedStudents.add(sid);
        deleted++;
      }
    }
  }
  
  return {ok: true, count: deleted};
}

function apiStudentFees(p) {
  const bills = readRows_(SHEET_NAMES.BILLS).filter(b => b.StudentID === p.studentId).map(b => ({
    billId: b.BillID, 
    amount: Number(b.Amount) || 0, 
    discount: Number(b.Discount) || 0, 
    paid: Number(b.Paid) || 0, 
    dueDate: b.DueDate, 
    status: b.Status, 
    note: b.Note
  })); 
  
  const totalDue = bills.reduce((sum, b) => sum + Math.max(0, b.amount - b.discount - b.paid), 0);
  return {ok: true, bills: bills, totalDue: totalDue};
}

function apiBatchDues(p) {
  let students = readRows_(SHEET_NAMES.STUDENTS);
  if (p.batchId && p.batchId.trim() !== '') {
    students = students.filter(s => s.BatchID === p.batchId);
  }
  
  const batches = readRows_(SHEET_NAMES.BATCHES); 
  const allBills = readRows_(SHEET_NAMES.BILLS);
  
  const list = students.map(st => {
    const b = batches.find(x => x.BatchID === st.BatchID);
    const due = allBills.filter(bill => bill.StudentID === st.StudentID).reduce((s, bill) => {
      return s + Math.max(0, (Number(bill.Amount) || 0) - (Number(bill.Discount) || 0) - (Number(bill.Paid) || 0));
    }, 0);
    
    return { 
      studentId: st.StudentID, 
      name: st.Name, 
      parentPhone: st.ParentPhone, 
      batchName: b ? b.BatchName : st.BatchID, 
      due: due 
    };
  }).filter(x => x.due > 0); 
  
  const grandTotal = list.reduce((s, x) => s + x.due, 0);
  return {ok: true, students: list, totalDue: grandTotal};
}

function apiCollectMultiFee(p) {
  const billIds = JSON.parse(p.billIds); 
  let remainingAmt = Number(p.amount);
  
  const sh = sheet_(SHEET_NAMES.BILLS); 
  const data = sh.getDataRange().getValues(); 
  const head = data[0];
  
  for (let b of billIds) {
    if (remainingAmt <= 0) break;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][head.indexOf('BillID')] === b) {
        const amt = Number(data[i][head.indexOf('Amount')]) || 0; 
        const disc = Number(data[i][head.indexOf('Discount')]) || 0; 
        const paid = Number(data[i][head.indexOf('Paid')]) || 0;
        
        const pending = amt - disc - paid;
        if (pending > 0) {
          const payNow = Math.min(remainingAmt, pending); 
          const newPaid = paid + payNow;
          
          sh.getRange(i + 1, head.indexOf('Paid') + 1).setValue(newPaid);
          sh.getRange(i + 1, head.indexOf('Status') + 1).setValue((newPaid + disc >= amt) ? 'PAID' : 'PARTIAL');
          
          const payId = 'PAY-' + Utilities.formatDate(now_(), 'Asia/Kolkata', 'yyMMddHHmmss') + randomLetters_(2);
          
          append_(SHEET_NAMES.PAYMENTS, { 
            PaymentID: payId, 
            StudentID: p.studentId, 
            BillID: b, 
            Amount: payNow, 
            PaidOn: istString_(now_()).slice(0, 10), 
            Mode: p.mode || 'CASH', 
            Note: p.note || '', 
            CreatedAt: istString_(now_()) 
          });
          
          remainingAmt -= payNow;
        } 
        break;
      }
    }
  } 
  return {ok: true};
}

function apiApplyMultiDiscount(p) {
  const billIds = JSON.parse(p.billIds); 
  let remainingAmt = Number(p.amount);
  
  const sh = sheet_(SHEET_NAMES.BILLS); 
  const data = sh.getDataRange().getValues(); 
  const head = data[0];
  
  for (let b of billIds) {
    if (remainingAmt <= 0) break;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][head.indexOf('BillID')] === b) {
        const amt = Number(data[i][head.indexOf('Amount')]) || 0; 
        const disc = Number(data[i][head.indexOf('Discount')]) || 0; 
        const paid = Number(data[i][head.indexOf('Paid')]) || 0;
        
        const pending = amt - disc - paid;
        if (pending > 0) {
          const discNow = Math.min(remainingAmt, pending); 
          const newDisc = disc + discNow;
          
          sh.getRange(i + 1, head.indexOf('Discount') + 1).setValue(newDisc);
          sh.getRange(i + 1, head.indexOf('Status') + 1).setValue((paid + newDisc >= amt) ? 'PAID' : 'PARTIAL');
          
          const oldNote = data[i][head.indexOf('Note')];
          const newNote = (oldNote ? oldNote + ' | ' : '') + (p.note || 'Discount');
          sh.getRange(i + 1, head.indexOf('Note') + 1).setValue(newNote);
          
          const did = 'DIS-' + Utilities.formatDate(now_(), 'Asia/Kolkata', 'yyMMddHHmmss') + randomLetters_(2);
          
          append_(SHEET_NAMES.DISCOUNTS, { 
            DiscountID: did, 
            StudentID: p.studentId, 
            BillID: b, 
            Amount: discNow, 
            Note: p.note || '', 
            CreatedAt: istString_(now_()) 
          });
          
          remainingAmt -= discNow;
        } 
        break;
      }
    }
  } 
  return {ok: true};
}

function apiAddPrevDue(p) {
  const bid = 'PREV-' + Utilities.formatDate(now_(), 'Asia/Kolkata', 'yyMMddHHmmss') + '-' + p.studentId;
  append_(SHEET_NAMES.BILLS, { 
    BillID: bid, 
    StudentID: p.studentId, 
    BatchID: '', 
    Month: 0, 
    Year: 0, 
    Amount: Number(p.amount || 0), 
    DueDate: istString_(now_()).slice(0, 10), 
    Discount: 0, 
    Paid: 0, 
    Status: 'UNPAID', 
    Note: 'Previous Due', 
    CreatedAt: istString_(now_()) 
  });
  return {ok: true};
}

/* ---------------- REPORTS ---------------- */
function apiCollectionReport(p) {
  const d = Number(p.days || 30); 
  const cutoff = new Date(now_().getTime() - d * 86400000);
  
  const students = readRows_(SHEET_NAMES.STUDENTS); 
  const batches = readRows_(SHEET_NAMES.BATCHES);
  
  const rows = readRows_(SHEET_NAMES.PAYMENTS).filter(x => new Date(x.PaidOn) >= cutoff).map(x => {
    const st = students.find(s => s.StudentID === x.StudentID) || {};
    
    // Filter by Batch
    if (p.batchId && p.batchId.trim() !== '' && st.BatchID !== p.batchId) return null;
    
    const b = batches.find(batch => batch.BatchID === st.BatchID) || {};
    return { 
      paidOn: x.PaidOn, 
      studentId: x.StudentID, 
      name: st.Name || '', 
      batchName: b.BatchName || st.BatchID || 'N/A', 
      amount: Number(x.Amount) || 0, 
      mode: x.Mode || '' 
    };
  }).filter(Boolean).sort((a, b) => a.paidOn < b.paidOn ? 1 : -1);
  
  const grandTotal = rows.reduce((s, r) => s + r.amount, 0);
  return {ok: true, rows: rows, total: grandTotal};
}
