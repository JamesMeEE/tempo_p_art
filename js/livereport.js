var _lrDateFrom = null;
var _lrDateTo = null;

function initLiveReportDateFilter() {
  var today = getTodayLocalStr();
  var elFrom = document.getElementById('lrDateFrom');
  var elTo = document.getElementById('lrDateTo');
  if (elFrom && !elFrom.value) elFrom.value = today;
  if (elTo && !elTo.value) elTo.value = today;
  _lrDateFrom = today;
  _lrDateTo = today;
}

function filterLiveReport() {
  var f = document.getElementById('lrDateFrom').value;
  var t = document.getElementById('lrDateTo').value;
  if (f && !t) t = f;
  if (t && !f) f = t;
  _lrDateFrom = f || null;
  _lrDateTo = t || null;
  loadLiveReport();
}

function resetLiveReportFilter() {
  var today = getTodayLocalStr();
  document.getElementById('lrDateFrom').value = today;
  document.getElementById('lrDateTo').value = today;
  _lrDateFrom = today;
  _lrDateTo = today;
  loadLiveReport();
}

async function loadLiveReport() {
  if (!currentUser || (currentUser.role !== 'Admin' && !isManager())) return;
  initLiveReportDateFilter();

  var dateFrom = _lrDateFrom;
  var dateTo = _lrDateTo;

  try {
    var dbData = await fetchSheetData('_database!A1:M31');
    var sellsData = await fetchSheetData('Sells!A:M');
    var tradeinsData = await fetchSheetData('Tradeins!A:O');
    var exchangesData = await fetchSheetData('Exchanges!A:T');
    var switchesData = await fetchSheetData('Switches!A:N');
    var freeExData = await fetchSheetData('FreeExchanges!A:J');
    var buybacksData = await fetchSheetData('Buybacks!A:L');
    var withdrawsData = await fetchSheetData('Withdraws!A:L');
    var closeData = await fetchSheetData('Close!A:K');

    var users = [];
    if (dbData && dbData.length > 33) {
      for (var i = 32; i < dbData.length; i++) {
        if (dbData[i] && dbData[i][2] && String(dbData[i][2]).trim()) {
          var role = String(dbData[i][0] || '').trim();
          var nickname = String(dbData[i][1] || '').trim();
          if (role === 'User' && nickname) {
            users.push(nickname);
          }
        }
      }
    }

    var salesUserData = {};
    for (var u = 0; u < users.length; u++) {
      var un = users[u];
      try {
        var ud = await fetchSheetData(un + '!A:I');
        var gd = await fetchSheetData(un + '_Gold!A:F');
        salesUserData[un] = { sheet: ud || [], gold: gd || [] };
      } catch(e) {
        salesUserData[un] = { sheet: [], gold: [] };
      }
    }

    renderSalesStatus(users, salesUserData, closeData, sellsData, tradeinsData, exchangesData, switchesData, freeExData, buybacksData, withdrawsData, dateFrom, dateTo);
    renderLRSummaryBoxes(sellsData, tradeinsData, exchangesData, switchesData, freeExData, buybacksData, withdrawsData, dateFrom, dateTo);
    renderLRPaymentSummary('lrSalesPayments', 'ยอดเงินที่ได้รับจากการขาย', ['SELL', 'TRADEIN', 'EXCHANGE', 'SWITCH', 'FREE_EXCHANGE', 'FREE-EX', 'WITHDRAW'], users, salesUserData, dateFrom, dateTo);
    renderLRPaymentSummary('lrBuybackPayments', 'ยอดเงินที่จ่าย Buyback', ['BUYBACK'], users, salesUserData, dateFrom, dateTo);
    renderLRStockSummary(sellsData, tradeinsData, exchangesData, switchesData, freeExData, buybacksData, withdrawsData, dateFrom, dateTo);
    renderLRGoldTable(sellsData, tradeinsData, exchangesData, switchesData, freeExData, buybacksData, withdrawsData, dateFrom, dateTo);
  } catch(e) {
    console.error('loadLiveReport error:', e);
  }
}

function lrInRange(dateVal, dateFrom, dateTo) {
  if (!dateFrom && !dateTo) return true;
  try {
    var d = new Date(dateVal);
    if (isNaN(d.getTime())) return false;
    var local = new Date(d.getTime() + 7 * 60 * 60000);
    var ds = local.toISOString().split('T')[0];
    if (dateFrom && ds < dateFrom) return false;
    if (dateTo && ds > dateTo) return false;
    return true;
  } catch(e) { return false; }
}

function renderSalesStatus(users, salesUserData, closeData, sellsData, tradeinsData, exchangesData, switchesData, freeExData, buybacksData, withdrawsData, dateFrom, dateTo) {
  var container = document.getElementById('lrSalesStatus');
  if (!container) return;
  var html = '';
  var weights = { 'G01': 150, 'G02': 75, 'G03': 30, 'G04': 15, 'G05': 7.5, 'G06': 3.75, 'G07': 1 };

  for (var u = 0; u < users.length; u++) {
    var name = users[u];
    var ud = salesUserData[name];
    var isOpen = ud.sheet.length > 1 && ud.sheet[1] && ud.sheet[1][0] && String(ud.sheet[1][0]).trim() !== '';

    var shiftClosed = false;
    if (closeData && closeData.length > 1) {
      for (var ci = 1; ci < closeData.length; ci++) {
        var cu = String(closeData[ci][1] || '').trim();
        var cs = String(closeData[ci][8] || '').trim();
        if (cu !== name) continue;
        if (cs === 'PENDING' || cs === 'APPROVED') {
          try {
            var cd = new Date(closeData[ci][2]);
            var cl = new Date(cd.getTime() + 7 * 60 * 60000).toISOString().split('T')[0];
            var today = getTodayLocalStr();
            if (cl === today) { shiftClosed = true; break; }
          } catch(e) {}
        }
      }
    }

    var statusText = shiftClosed ? '🔴 ปิดกะแล้ว' : (isOpen ? '🟢 เปิดกะอยู่' : '⚪ ยังไม่เปิดกะ');
    var statusColor = shiftClosed ? '#f44336' : (isOpen ? '#4caf50' : '#888');

    var sellCount = 0, sellLAK = 0, sellG = 0;
    var bbCount = 0, bbLAK = 0, bbG = 0;
    var wdCount = 0, wdLAK = 0, wdG = 0;

    if (sellsData && sellsData.length > 1) {
      for (var si = 1; si < sellsData.length; si++) {
        if (String(sellsData[si][11] || '') !== name) continue;
        if (String(sellsData[si][10] || '') !== 'COMPLETED') continue;
        if (!lrInRange(sellsData[si][9], dateFrom, dateTo)) continue;
        sellCount++;
        sellLAK += parseFloat(String(sellsData[si][3]).replace(/,/g, '')) || 0;
        try { var items = JSON.parse(sellsData[si][2]); items.forEach(function(it) { sellG += (weights[it.productId] || 0) * it.qty; }); } catch(e) {}
      }
    }
    [tradeinsData, exchangesData].forEach(function(sheet) {
      if (!sheet || sheet.length <= 1) return;
      for (var ti = 1; ti < sheet.length; ti++) {
        if (String(sheet[ti][13] || '') !== name) continue;
        if (String(sheet[ti][12] || '') !== 'COMPLETED') continue;
        if (!lrInRange(sheet[ti][11], dateFrom, dateTo)) continue;
        sellCount++;
        sellLAK += parseFloat(String(sheet[ti][6]).replace(/,/g, '')) || 0;
        try { var nit = JSON.parse(sheet[ti][3]); nit.forEach(function(it) { sellG += (weights[it.productId] || 0) * it.qty; }); } catch(e) {}
      }
    });
    if (switchesData && switchesData.length > 1) {
      for (var swi = 1; swi < switchesData.length; swi++) {
        if (String(switchesData[swi][13] || '') !== name) continue;
        if (String(switchesData[swi][12] || '') !== 'COMPLETED') continue;
        if (!lrInRange(switchesData[swi][11], dateFrom, dateTo)) continue;
        sellCount++;
        sellLAK += parseFloat(String(switchesData[swi][6]).replace(/,/g, '')) || 0;
        try { var nit2 = JSON.parse(switchesData[swi][3]); nit2.forEach(function(it) { sellG += (weights[it.productId] || 0) * it.qty; }); } catch(e) {}
      }
    }
    if (freeExData && freeExData.length > 1) {
      for (var fi = 1; fi < freeExData.length; fi++) {
        if (String(freeExData[fi][9] || '') !== name) continue;
        if (String(freeExData[fi][8] || '') !== 'COMPLETED') continue;
        if (!lrInRange(freeExData[fi][7], dateFrom, dateTo)) continue;
        sellCount++;
        sellLAK += parseFloat(String(freeExData[fi][5]).replace(/,/g, '')) || 0;
        try { var nit3 = JSON.parse(freeExData[fi][3]); nit3.forEach(function(it) { sellG += (weights[it.productId] || 0) * it.qty; }); } catch(e) {}
      }
    }
    if (withdrawsData && withdrawsData.length > 1) {
      for (var wi = 1; wi < withdrawsData.length; wi++) {
        if (String(withdrawsData[wi][8] || '') !== name) continue;
        if (String(withdrawsData[wi][7] || '') !== 'COMPLETED') continue;
        if (!lrInRange(withdrawsData[wi][6], dateFrom, dateTo)) continue;
        wdCount++;
        wdLAK += parseFloat(String(withdrawsData[wi][4]).replace(/,/g, '')) || 0;
        try { var wit = JSON.parse(withdrawsData[wi][2]); wit.forEach(function(it) { wdG += (weights[it.productId] || 0) * it.qty; }); } catch(e) {}
      }
    }
    if (buybacksData && buybacksData.length > 1) {
      for (var bi = 1; bi < buybacksData.length; bi++) {
        var bbCreator = String(buybacksData[bi][11] || '').trim();
        if (bbCreator !== name) continue;
        var bbStatus = String(buybacksData[bi][10] || '').trim();
        if (bbStatus !== 'COMPLETED' && bbStatus !== 'PARTIAL') continue;
        if (!lrInRange(buybacksData[bi][9], dateFrom, dateTo)) continue;
        bbCount++;
        bbLAK += parseFloat(String(buybacksData[bi][6]).replace(/,/g, '')) || 0;
        try { var bbit = JSON.parse(buybacksData[bi][2]); bbit.forEach(function(it) { bbG += (weights[it.productId] || 0) * it.qty; }); } catch(e) {}
      }
    }

    var cashLAK = 0, cashTHB = 0, cashUSD = 0;
    var oldGoldG = 0;
    if (isOpen && ud.sheet.length > 1) {
      for (var r = 1; r < ud.sheet.length; r++) {
        if (String(ud.sheet[r][4] || '').trim() === 'Cash') {
          var cur = String(ud.sheet[r][3] || '').trim();
          var amt = parseFloat(ud.sheet[r][2]) || 0;
          if (cur === 'LAK') cashLAK += amt;
          else if (cur === 'THB') cashTHB += amt;
          else if (cur === 'USD') cashUSD += amt;
        }
      }
    }
    if (ud.gold.length > 1) {
      for (var gi = 1; gi < ud.gold.length; gi++) {
        var pid = String(ud.gold[gi][0] || '').trim();
        var qty = parseFloat(ud.gold[gi][1]) || 0;
        oldGoldG += (weights[pid] || 0) * qty;
      }
    }

    html += '<div style="background:var(--bg-secondary);border:1px solid var(--border-color);border-radius:12px;padding:16px;margin-bottom:12px;">';
    html += '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;">';
    html += '<span style="font-weight:700;font-size:16px;color:var(--gold-primary);">' + name + '</span>';
    html += '<span style="font-size:13px;color:' + statusColor + ';font-weight:600;">' + statusText + '</span>';
    html += '</div>';
    if (isOpen) {
      html += '<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:10px;font-size:13px;">';
      html += '<div><span style="color:var(--text-secondary);">Sales:</span> ' + formatNumber(sellLAK) + ' LAK | ' + sellG.toFixed(2) + 'g | ' + sellCount + ' บิล</div>';
      html += '<div><span style="color:var(--text-secondary);">Withdraw:</span> ' + formatNumber(wdLAK) + ' LAK | ' + wdG.toFixed(2) + 'g | ' + wdCount + ' บิล</div>';
      html += '<div><span style="color:var(--text-secondary);">Buyback:</span> ' + formatNumber(bbLAK) + ' LAK | ' + bbG.toFixed(2) + 'g | ' + bbCount + ' บิล</div>';
      html += '<div><span style="color:var(--text-secondary);">เงินสด:</span> ' + formatNumber(cashLAK) + ' LAK | ' + formatNumber(cashTHB) + ' THB | ' + formatNumber(cashUSD) + ' USD</div>';
      html += '<div><span style="color:var(--text-secondary);">ทองเก่า:</span> ' + oldGoldG.toFixed(2) + ' g</div>';
      html += '</div>';
    }
    html += '</div>';
  }

  container.innerHTML = html || '<div style="text-align:center;color:var(--text-secondary);padding:20px;">ไม่พบข้อมูล Sales</div>';
}

function renderLRSummaryBoxes(sellsData, tradeinsData, exchangesData, switchesData, freeExData, buybacksData, withdrawsData, dateFrom, dateTo) {
  var weights = { 'G01': 150, 'G02': 75, 'G03': 30, 'G04': 15, 'G05': 7.5, 'G06': 3.75, 'G07': 1 };
  var salesLAK = 0, salesG = 0, salesCount = 0;
  var bbLAK = 0, bbG = 0, bbCount = 0;
  var wdLAK = 0, wdG = 0, wdCount = 0;

  function sumSheet(data, totalCol, itemsCol, statusCol, dateCol, createdCol, statuses) {
    var lak = 0, g = 0, count = 0;
    if (!data || data.length <= 1) return { lak: 0, g: 0, count: 0 };
    for (var i = 1; i < data.length; i++) {
      var st = String(data[i][statusCol] || '').trim();
      if (statuses.indexOf(st) === -1) continue;
      if (!lrInRange(data[i][dateCol], dateFrom, dateTo)) continue;
      count++;
      lak += parseFloat(String(data[i][totalCol]).replace(/,/g, '')) || 0;
      try { var it = JSON.parse(data[i][itemsCol]); it.forEach(function(x) { g += (weights[x.productId] || 0) * x.qty; }); } catch(e) {}
    }
    return { lak: lak, g: g, count: count };
  }

  var s1 = sumSheet(sellsData, 3, 2, 10, 9, 11, ['COMPLETED']);
  salesLAK += s1.lak; salesG += s1.g; salesCount += s1.count;
  var s2 = sumSheet(tradeinsData, 6, 3, 12, 11, 13, ['COMPLETED']);
  salesLAK += s2.lak; salesG += s2.g; salesCount += s2.count;
  var s3 = sumSheet(exchangesData, 6, 3, 12, 11, 13, ['COMPLETED']);
  salesLAK += s3.lak; salesG += s3.g; salesCount += s3.count;
  var s4 = sumSheet(switchesData, 6, 3, 12, 11, 13, ['COMPLETED']);
  salesLAK += s4.lak; salesG += s4.g; salesCount += s4.count;
  var s5 = sumSheet(freeExData, 5, 3, 8, 7, 9, ['COMPLETED']);
  salesLAK += s5.lak; salesG += s5.g; salesCount += s5.count;

  var b1 = sumSheet(buybacksData, 6, 2, 10, 9, 11, ['COMPLETED', 'PARTIAL']);
  bbLAK = b1.lak; bbG = b1.g; bbCount = b1.count;

  var w1 = sumSheet(withdrawsData, 4, 2, 7, 6, 8, ['COMPLETED']);
  wdLAK = w1.lak; wdG = w1.g; wdCount = w1.count;

  var netLAK = salesLAK - bbLAK;
  var netG = salesG - bbG;

  var boxStyle = 'background:var(--bg-secondary);border:1px solid var(--border-color);border-radius:12px;padding:20px;text-align:center;';

  var html = '<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:15px;margin-bottom:25px;">';
  html += '<div style="' + boxStyle + '"><div style="color:var(--text-secondary);font-size:12px;margin-bottom:8px;">NET SELL</div><div style="font-size:22px;font-weight:700;color:var(--gold-primary);">' + formatNumber(netLAK) + '</div><div style="font-size:12px;color:var(--text-secondary);margin-top:4px;">' + netG.toFixed(2) + ' g</div></div>';
  html += '<div style="' + boxStyle + '"><div style="color:var(--text-secondary);font-size:12px;margin-bottom:8px;">SALES</div><div style="font-size:22px;font-weight:700;color:#4caf50;">' + formatNumber(salesLAK) + '</div><div style="font-size:12px;color:var(--text-secondary);margin-top:4px;">' + salesG.toFixed(2) + ' g | ' + salesCount + ' บิล</div></div>';
  html += '<div style="' + boxStyle + '"><div style="color:var(--text-secondary);font-size:12px;margin-bottom:8px;">BUYBACK</div><div style="font-size:22px;font-weight:700;color:#f44336;">' + formatNumber(bbLAK) + '</div><div style="font-size:12px;color:var(--text-secondary);margin-top:4px;">' + bbG.toFixed(2) + ' g | ' + bbCount + ' บิล</div></div>';
  html += '<div style="' + boxStyle + '"><div style="color:var(--text-secondary);font-size:12px;margin-bottom:8px;">WITHDRAW</div><div style="font-size:22px;font-weight:700;color:#ff9800;">' + formatNumber(wdLAK) + '</div><div style="font-size:12px;color:var(--text-secondary);margin-top:4px;">' + wdG.toFixed(2) + ' g | ' + wdCount + ' บิล</div></div>';
  html += '</div>';

  document.getElementById('lrSummaryBoxes').innerHTML = html;
}

function renderLRPaymentSummary(containerId, title, types, users, salesUserData, dateFrom, dateTo) {
  var container = document.getElementById(containerId);
  if (!container) return;

  var methods = ['Cash', 'BCEL', 'LDB', 'Other'];
  var currencies = ['LAK', 'THB', 'USD'];
  var totals = {};
  methods.forEach(function(m) {
    totals[m] = {};
    currencies.forEach(function(c) { totals[m][c] = 0; });
  });

  for (var u = 0; u < users.length; u++) {
    var ud = salesUserData[users[u]];
    if (!ud || !ud.sheet || ud.sheet.length <= 1) continue;
    for (var r = 1; r < ud.sheet.length; r++) {
      var row = ud.sheet[r];
      var rType = String(row[1] || '').trim();
      var baseType = rType.replace('_CHANGE', '');
      var matched = false;
      for (var t = 0; t < types.length; t++) {
        if (baseType === types[t]) { matched = true; break; }
      }
      if (!matched) continue;
      if (!lrInRange(row[7], dateFrom, dateTo)) continue;
      var amt = parseFloat(row[2]) || 0;
      var cur = String(row[3] || '').trim();
      var method = String(row[4] || '').trim();
      var bank = String(row[5] || '').trim();

      var key = 'Cash';
      if (method === 'Bank') {
        if (bank === 'BCEL') key = 'BCEL';
        else if (bank === 'LDB') key = 'LDB';
        else key = 'Other';
      }
      if (currencies.indexOf(cur) >= 0 && totals[key]) {
        totals[key][cur] += amt;
      }
    }
  }

  var thStyle = 'background:#2d2d2d;color:#d4af37;border:1px solid rgba(212,175,55,0.5);padding:10px 8px;font-size:12px;text-align:center;font-weight:700;';
  var tdStyle = 'border:1px solid var(--border-color);padding:8px;text-align:right;font-size:13px;';

  var html = '<div style="background:var(--bg-secondary);border:1px solid var(--border-color);border-radius:12px;padding:16px;margin-bottom:20px;">';
  html += '<h3 style="color:var(--gold-primary);font-size:16px;margin-bottom:12px;">' + title + '</h3>';
  html += '<div class="table-container"><table><thead><tr>';
  html += '<th style="' + thStyle + '">ช่องทาง</th>';
  currencies.forEach(function(c) { html += '<th style="' + thStyle + '">' + c + '</th>'; });
  html += '</tr></thead><tbody>';

  var grandTotal = {};
  currencies.forEach(function(c) { grandTotal[c] = 0; });

  methods.forEach(function(m) {
    html += '<tr>';
    html += '<td style="' + tdStyle + 'text-align:left;font-weight:600;">' + m + '</td>';
    currencies.forEach(function(c) {
      var val = totals[m][c];
      grandTotal[c] += val;
      var color = val > 0 ? '#4caf50' : (val < 0 ? '#f44336' : 'var(--text-secondary)');
      html += '<td style="' + tdStyle + 'color:' + color + ';">' + formatNumber(Math.round(val)) + '</td>';
    });
    html += '</tr>';
  });

  html += '<tr style="background:rgba(212,175,55,0.1);font-weight:700;">';
  html += '<td style="' + tdStyle + 'text-align:left;color:var(--gold-primary);">รวม</td>';
  currencies.forEach(function(c) {
    var val = grandTotal[c];
    var color = val > 0 ? '#4caf50' : (val < 0 ? '#f44336' : 'var(--text-secondary)');
    html += '<td style="' + tdStyle + 'color:' + color + ';font-weight:700;">' + formatNumber(Math.round(val)) + '</td>';
  });
  html += '</tr>';
  html += '</tbody></table></div></div>';

  container.innerHTML = html;
}

function renderLRStockSummary(sellsData, tradeinsData, exchangesData, switchesData, freeExData, buybacksData, withdrawsData, dateFrom, dateTo) {
  var container = document.getElementById('lrStockSummary');
  if (!container) return;
  var weights = { 'G01': 150, 'G02': 75, 'G03': 30, 'G04': 15, 'G05': 7.5, 'G06': 3.75, 'G07': 1 };
  var newOutG = 0, oldInG = 0;

  function sumNewOut(data, itemsCol, statusCol, dateCol, statuses) {
    if (!data || data.length <= 1) return 0;
    var g = 0;
    for (var i = 1; i < data.length; i++) {
      if (statuses.indexOf(String(data[i][statusCol] || '').trim()) === -1) continue;
      if (!lrInRange(data[i][dateCol], dateFrom, dateTo)) continue;
      try { var it = JSON.parse(data[i][itemsCol]); it.forEach(function(x) { g += (weights[x.productId] || 0) * x.qty; }); } catch(e) {}
    }
    return g;
  }

  newOutG += sumNewOut(sellsData, 2, 10, 9, ['COMPLETED']);
  newOutG += sumNewOut(tradeinsData, 3, 12, 11, ['COMPLETED']);
  newOutG += sumNewOut(exchangesData, 3, 12, 11, ['COMPLETED']);
  newOutG += sumNewOut(switchesData, 3, 12, 11, ['COMPLETED']);
  newOutG += sumNewOut(freeExData, 3, 8, 7, ['COMPLETED']);
  newOutG += sumNewOut(withdrawsData, 2, 7, 6, ['COMPLETED']);

  function sumOldIn(data, itemsCol, statusCol, dateCol, statuses) {
    if (!data || data.length <= 1) return 0;
    var g = 0;
    for (var i = 1; i < data.length; i++) {
      if (statuses.indexOf(String(data[i][statusCol] || '').trim()) === -1) continue;
      if (!lrInRange(data[i][dateCol], dateFrom, dateTo)) continue;
      try { var it = JSON.parse(data[i][itemsCol]); it.forEach(function(x) { g += (weights[x.productId] || 0) * x.qty; }); } catch(e) {}
    }
    return g;
  }
  oldInG += sumOldIn(tradeinsData, 2, 12, 11, ['COMPLETED']);
  oldInG += sumOldIn(exchangesData, 2, 12, 11, ['COMPLETED']);
  if (exchangesData && exchangesData.length > 1) {
    for (var ei = 1; ei < exchangesData.length; ei++) {
      if (String(exchangesData[ei][12] || '').trim() !== 'COMPLETED') continue;
      if (!lrInRange(exchangesData[ei][11], dateFrom, dateTo)) continue;
      try { if (exchangesData[ei][14]) { var swOld = JSON.parse(exchangesData[ei][14]); swOld.forEach(function(x) { oldInG += (weights[x.productId] || 0) * x.qty; }); } } catch(e) {}
      try { if (exchangesData[ei][16]) { var feOld = JSON.parse(exchangesData[ei][16]); feOld.forEach(function(x) { oldInG += (weights[x.productId] || 0) * x.qty; }); } } catch(e) {}
    }
  }
  oldInG += sumOldIn(switchesData, 2, 12, 11, ['COMPLETED']);
  oldInG += sumOldIn(freeExData, 2, 8, 7, ['COMPLETED']);
  oldInG += sumOldIn(buybacksData, 2, 10, 9, ['COMPLETED', 'PARTIAL']);

  var boxStyle = 'background:var(--bg-secondary);border:1px solid var(--border-color);border-radius:12px;padding:20px;text-align:center;';
  var html = '<div style="display:grid;grid-template-columns:1fr 1fr;gap:15px;margin-bottom:20px;">';
  html += '<div style="' + boxStyle + '"><div style="color:var(--text-secondary);font-size:12px;margin-bottom:8px;">NEW OUT (ทองใหม่ออก)</div><div style="font-size:24px;font-weight:700;color:#f44336;">' + newOutG.toFixed(2) + ' g</div></div>';
  html += '<div style="' + boxStyle + '"><div style="color:var(--text-secondary);font-size:12px;margin-bottom:8px;">OLD IN (ทองเก่ารับเข้า)</div><div style="font-size:24px;font-weight:700;color:#4caf50;">' + oldInG.toFixed(2) + ' g</div></div>';
  html += '</div>';
  container.innerHTML = html;
}

function renderLRGoldTable(sellsData, tradeinsData, exchangesData, switchesData, freeExData, buybacksData, withdrawsData, dateFrom, dateTo) {
  var container = document.getElementById('lrGoldTable');
  if (!container) return;
  var products = ['G01', 'G02', 'G03', 'G04', 'G05', 'G06', 'G07'];
  var names = { 'G01': '10 บาท', 'G02': '5 บาท', 'G03': '2 บาท', 'G04': '1 บาท', 'G05': '2 สลึง', 'G06': '1 สลึง', 'G07': '1 กรัม' };
  var out = {}, inn = {};
  products.forEach(function(p) { out[p] = 0; inn[p] = 0; });

  function countOut(data, itemsCol, statusCol, dateCol, statuses) {
    if (!data || data.length <= 1) return;
    for (var i = 1; i < data.length; i++) {
      if (statuses.indexOf(String(data[i][statusCol] || '').trim()) === -1) continue;
      if (!lrInRange(data[i][dateCol], dateFrom, dateTo)) continue;
      try { var it = JSON.parse(data[i][itemsCol]); it.forEach(function(x) { if (out[x.productId] !== undefined) out[x.productId] += x.qty; }); } catch(e) {}
    }
  }
  function countIn(data, itemsCol, statusCol, dateCol, statuses) {
    if (!data || data.length <= 1) return;
    for (var i = 1; i < data.length; i++) {
      if (statuses.indexOf(String(data[i][statusCol] || '').trim()) === -1) continue;
      if (!lrInRange(data[i][dateCol], dateFrom, dateTo)) continue;
      try { var it = JSON.parse(data[i][itemsCol]); it.forEach(function(x) { if (inn[x.productId] !== undefined) inn[x.productId] += x.qty; }); } catch(e) {}
    }
  }

  countOut(sellsData, 2, 10, 9, ['COMPLETED']);
  countOut(tradeinsData, 3, 12, 11, ['COMPLETED']);
  countOut(exchangesData, 3, 12, 11, ['COMPLETED']);
  countOut(switchesData, 3, 12, 11, ['COMPLETED']);
  countOut(freeExData, 3, 8, 7, ['COMPLETED']);
  countOut(withdrawsData, 2, 7, 6, ['COMPLETED']);

  countIn(tradeinsData, 2, 12, 11, ['COMPLETED']);
  countIn(exchangesData, 2, 12, 11, ['COMPLETED']);
  countIn(switchesData, 2, 12, 11, ['COMPLETED']);
  countIn(freeExData, 2, 8, 7, ['COMPLETED']);
  countIn(buybacksData, 2, 10, 9, ['COMPLETED', 'PARTIAL']);
  if (exchangesData && exchangesData.length > 1) {
    for (var ei = 1; ei < exchangesData.length; ei++) {
      if (String(exchangesData[ei][12] || '').trim() !== 'COMPLETED') continue;
      if (!lrInRange(exchangesData[ei][11], dateFrom, dateTo)) continue;
      try { if (exchangesData[ei][14]) { var sw = JSON.parse(exchangesData[ei][14]); sw.forEach(function(x) { if (inn[x.productId] !== undefined) inn[x.productId] += x.qty; }); } } catch(e) {}
      try { if (exchangesData[ei][16]) { var fe = JSON.parse(exchangesData[ei][16]); fe.forEach(function(x) { if (inn[x.productId] !== undefined) inn[x.productId] += x.qty; }); } } catch(e) {}
    }
  }

  var thStyle = 'background:#2d2d2d;color:#d4af37;border:1px solid rgba(212,175,55,0.5);padding:10px 8px;font-size:12px;text-align:center;font-weight:700;';
  var tdStyle = 'border:1px solid var(--border-color);padding:8px;text-align:center;font-size:13px;';

  var html = '<div style="background:var(--bg-secondary);border:1px solid var(--border-color);border-radius:12px;padding:16px;">';
  html += '<h3 style="color:var(--gold-primary);font-size:16px;margin-bottom:12px;">รายละเอียดทองแต่ละ Product</h3>';
  html += '<div class="table-container"><table><thead><tr>';
  html += '<th style="' + thStyle + '">Product</th>';
  html += '<th style="' + thStyle + '">จ่ายออก (ชิ้น)</th>';
  html += '<th style="' + thStyle + '">ได้รับ (ชิ้น)</th>';
  html += '</tr></thead><tbody>';

  var totalOut = 0, totalIn = 0;
  products.forEach(function(p) {
    totalOut += out[p];
    totalIn += inn[p];
    html += '<tr>';
    html += '<td style="' + tdStyle + 'font-weight:600;">' + names[p] + ' (' + p + ')</td>';
    html += '<td style="' + tdStyle + 'color:' + (out[p] > 0 ? '#f44336' : 'var(--text-secondary)') + ';">' + out[p] + '</td>';
    html += '<td style="' + tdStyle + 'color:' + (inn[p] > 0 ? '#4caf50' : 'var(--text-secondary)') + ';">' + inn[p] + '</td>';
    html += '</tr>';
  });

  html += '<tr style="background:rgba(212,175,55,0.1);font-weight:700;">';
  html += '<td style="' + tdStyle + 'color:var(--gold-primary);font-weight:700;">รวม</td>';
  html += '<td style="' + tdStyle + 'color:#f44336;font-weight:700;">' + totalOut + '</td>';
  html += '<td style="' + tdStyle + 'color:#4caf50;font-weight:700;">' + totalIn + '</td>';
  html += '</tr>';
  html += '</tbody></table></div></div>';

  container.innerHTML = html;
}

async function loadSalesInfoBar() {
  if (!currentUser || currentUser.role !== 'User') {
    var bar = document.getElementById('salesInfoBar');
    if (bar) bar.style.display = 'none';
    return;
  }
  try {
    var bar = document.getElementById('salesInfoBar');
    if (bar) bar.style.display = 'block';

    var sheetName = currentUser.nickname;
    var userData = await fetchSheetData(sheetName + '!A:I');
    var goldData = await fetchSheetData(sheetName + '_Gold!A:F');
    var weights = { 'G01': 150, 'G02': 75, 'G03': 30, 'G04': 15, 'G05': 7.5, 'G06': 3.75, 'G07': 1 };

    var cashLAK = 0, cashTHB = 0, cashUSD = 0;
    if (userData && userData.length > 1) {
      for (var i = 1; i < userData.length; i++) {
        if (String(userData[i][4] || '').trim() === 'Cash') {
          var cur = String(userData[i][3] || '').trim();
          var amt = parseFloat(userData[i][2]) || 0;
          if (cur === 'LAK') cashLAK += amt;
          else if (cur === 'THB') cashTHB += amt;
          else if (cur === 'USD') cashUSD += amt;
        }
      }
    }

    var oldGoldG = 0;
    if (goldData && goldData.length > 1) {
      for (var gi = 1; gi < goldData.length; gi++) {
        var pid = String(goldData[gi][0] || '').trim();
        var qty = parseFloat(goldData[gi][1]) || 0;
        oldGoldG += (weights[pid] || 0) * qty;
      }
    }

    var s1b = currentPricing.sell1Baht || 0;
    var sellPrice = calculateSellPrice('G04', s1b);
    var buybackPrice = calculateBuybackPrice('G04', s1b);
    var exFee = EXCHANGE_FEES['G04'] || 0;
    var swFee = EXCHANGE_FEES_SWITCH['G04'] || 0;

    document.getElementById('siCashLAK').textContent = formatNumber(cashLAK) + ' LAK';
    document.getElementById('siCashTHB').textContent = formatNumber(cashTHB) + ' THB';
    document.getElementById('siCashUSD').textContent = formatNumber(cashUSD) + ' USD';
    document.getElementById('siOldGold').textContent = oldGoldG.toFixed(2) + ' g';
    document.getElementById('siSellPrice').textContent = formatNumber(sellPrice);
    document.getElementById('siBuybackPrice').textContent = formatNumber(buybackPrice);
    document.getElementById('siExFee').textContent = formatNumber(exFee);
    document.getElementById('siSwFee').textContent = formatNumber(swFee);
  } catch(e) {
    console.error('loadSalesInfoBar error:', e);
  }
}
