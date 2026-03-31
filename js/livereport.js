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
    var dbData = await fetchSheetData('_database!A1:M100');
    var sellsData = await fetchSheetData('Sells!A:M');
    var tradeinsData = await fetchSheetData('Tradeins!A:O');
    var exchangesData = await fetchSheetData('Exchanges!A:T');
    var buybacksData = await fetchSheetData('Buybacks!A:L');
    var withdrawsData = await fetchSheetData('Withdraws!A:L');
    var closeData = await fetchSheetData('Close!A:K');

    var users = [];
    if (dbData && dbData.length > 33) {
      for (var i = 33; i < dbData.length; i++) {
        if (dbData[i] && dbData[i][2] && String(dbData[i][2]).trim()) {
          var role = String(dbData[i][0] || '').trim();
          var nickname = String(dbData[i][1] || '').trim();
          if (role === 'Sales' && nickname) {
            users.push(nickname);
          }
        }
      }
    }

    var batchRanges = ['Switches!A:N', 'FreeExchanges!A:J', '_log_cashbank!A:I'];
    for (var u = 0; u < users.length; u++) {
      batchRanges.push(users[u] + '!A:I');
      batchRanges.push(users[u] + '_Gold!A:F');
    }

    var batchResult = {};
    try {
      var br = await callAppsScript('BATCH_READ', { ranges: JSON.stringify(batchRanges) });
      if (br && br.success && br.data) batchResult = br.data;
    } catch(e) {}

    var switchesData = batchResult['Switches!A:N'] || [];
    var freeExData = batchResult['FreeExchanges!A:J'] || [];
    var logCashbankData = batchResult['_log_cashbank!A:I'] || [];

    var salesUserData = {};
    for (var u = 0; u < users.length; u++) {
      var un = users[u];
      salesUserData[un] = {
        sheet: batchResult[un + '!A:I'] || [],
        gold: batchResult[un + '_Gold!A:F'] || []
      };
    }

    renderSalesStatus(users, salesUserData, closeData, logCashbankData, sellsData, tradeinsData, exchangesData, switchesData, freeExData, buybacksData, withdrawsData, dateFrom, dateTo);
    renderLRSummaryBoxes(sellsData, tradeinsData, exchangesData, switchesData, freeExData, buybacksData, withdrawsData, dateFrom, dateTo);
    renderLRPaymentSummary('lrSalesPayments', 'ยอดเงินที่ได้รับจากการขาย', ['SELL', 'TRADEIN', 'EXCHANGE', 'SWITCH', 'FREE_EXCHANGE', 'FREE-EX', 'WITHDRAW'], users, salesUserData, dateFrom, dateTo);
    renderLRPaymentSummary('lrBuybackPayments', 'ยอดเงินที่จ่าย Buyback', ['BUYBACK'], users, salesUserData, dateFrom, dateTo);
    renderLRStockSummary(sellsData, tradeinsData, exchangesData, switchesData, freeExData, buybacksData, withdrawsData, dateFrom, dateTo);
    var stockMoveNewData = await fetchSheetData('StockMove_New!A:K');
    var stockMoveOldData = await fetchSheetData('StockMove_Old!A:K');
    renderLRGoldTable(stockMoveNewData, stockMoveOldData, dateFrom, dateTo);
  } catch(e) {
    console.error('loadLiveReport error:', e);
  }
}

function lrInRange(dateVal, dateFrom, dateTo) {
  if (!dateFrom && !dateTo) return true;
  try {
    var d = parseSheetDate(dateVal);
    if (!d || isNaN(d.getTime())) return false;
    var ds = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
    if (dateFrom && ds < dateFrom) return false;
    if (dateTo && ds > dateTo) return false;
    return true;
  } catch(e) { return false; }
}

function renderSalesStatus(users, salesUserData, closeData, logCashbankData, sellsData, tradeinsData, exchangesData, switchesData, freeExData, buybacksData, withdrawsData, dateFrom, dateTo) {
  var container = document.getElementById('lrSalesStatus');
  if (!container) return;
  var html = '';
  var weights = { 'G01': 150, 'G02': 75, 'G03': 30, 'G04': 15, 'G05': 7.5, 'G06': 3.75, 'G07': 1 };

  for (var u = 0; u < users.length; u++) {
    var name = users[u];
    var ud = salesUserData[name];
    var isOpen = ud.sheet.length > 1 && ud.sheet[1] && ud.sheet[1][0] && String(ud.sheet[1][0]).trim() !== '';

    var shiftClosed = false;
    var closeRow = null;
    if (closeData && closeData.length > 1) {
      for (var ci = 1; ci < closeData.length; ci++) {
        var cu = String(closeData[ci][1] || '').trim();
        var cs = String(closeData[ci][8] || '').trim();
        if (cu !== name) continue;
        if (cs === 'PENDING' || cs === 'APPROVED') {
          try {
            var cd = parseSheetDate(closeData[ci][2]);
            if (cd) {
              var cl = cd.getFullYear() + '-' + String(cd.getMonth() + 1).padStart(2, '0') + '-' + String(cd.getDate()).padStart(2, '0');
              var today = getTodayLocalStr();
              if (cl === today) { shiftClosed = true; closeRow = closeData[ci]; break; }
            }
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
        if (bbStatus !== 'COMPLETED') continue;
        if (!lrInRange(buybacksData[bi][9], dateFrom, dateTo)) continue;
        bbCount++;
        bbLAK += parseFloat(String(buybacksData[bi][6]).replace(/,/g, '')) || 0;
        try { var bbit = JSON.parse(buybacksData[bi][2]); bbit.forEach(function(it) { bbG += (weights[it.productId] || 0) * it.qty; }); } catch(e) {}
      }
    }

    var cashLAK = 0, cashTHB = 0, cashUSD = 0;
    var oldGoldG = 0;

    if (shiftClosed) {
      if (closeRow) {
        cashLAK = parseFloat(closeRow[3]) || 0;
        cashTHB = parseFloat(closeRow[4]) || 0;
        cashUSD = parseFloat(closeRow[5]) || 0;
        try {
          var ogJson = closeRow[6];
          if (ogJson) {
            var ogItems = typeof ogJson === 'string' ? JSON.parse(ogJson) : ogJson;
            if (Array.isArray(ogItems)) {
              ogItems.forEach(function(it) { oldGoldG += (weights[it.productId] || 0) * (it.qty || 0); });
            }
          }
        } catch(e) {}
      }
    } else if (isOpen) {
      for (var r = 1; r < ud.sheet.length; r++) {
        if (String(ud.sheet[r][4] || '').trim() === 'Cash') {
          var cur = String(ud.sheet[r][3] || '').trim();
          var amt = parseFloat(ud.sheet[r][2]) || 0;
          if (cur === 'LAK') cashLAK += amt;
          else if (cur === 'THB') cashTHB += amt;
          else if (cur === 'USD') cashUSD += amt;
        }
      }
      if (ud.gold.length > 1) {
        for (var gi = 1; gi < ud.gold.length; gi++) {
          var pid = String(ud.gold[gi][0] || '').trim();
          var qty = parseFloat(ud.gold[gi][1]) || 0;
          oldGoldG += (weights[pid] || 0) * qty;
        }
      }
    }

    html += '<div style="background:var(--bg-secondary);border:1px solid var(--border-color);border-radius:12px;padding:16px;margin-bottom:12px;">';
    html += '<div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:' + (isOpen || shiftClosed ? '10' : '0') + 'px;">';
    html += '<span style="font-weight:700;font-size:16px;color:var(--gold-primary);">' + name + '</span>';
    html += '<span style="font-size:13px;color:' + statusColor + ';font-weight:600;">' + statusText + '</span>';
    html += '</div>';
    if (isOpen || shiftClosed) {
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

  function calcG(itemsJson) {
    var g = 0;
    try { var it = typeof itemsJson === 'string' ? JSON.parse(itemsJson) : itemsJson; it.forEach(function(x) { g += (weights[x.productId] || 0) * x.qty; }); } catch(e) {}
    return g;
  }

  function filterRows(data, statusCol, dateCol, statuses) {
    if (!data || data.length <= 1) return [];
    return data.slice(1).filter(function(r) {
      var st = String(r[statusCol] || '').trim();
      return statuses.indexOf(st) !== -1 && lrInRange(r[dateCol], dateFrom, dateTo);
    });
  }

  var sellRows = filterRows(sellsData, 10, 9, ['COMPLETED', 'PAID']);
  var tradeinRows = filterRows(tradeinsData, 12, 11, ['COMPLETED', 'PAID']);
  var exchangeRows = filterRows(exchangesData, 12, 11, ['COMPLETED', 'PAID']);
  var switchRows = filterRows(switchesData, 12, 11, ['COMPLETED', 'PAID']);
  var freeExRows = filterRows(freeExData, 8, 7, ['COMPLETED', 'PAID']);
  var buybackRows = filterRows(buybacksData, 10, 9, ['COMPLETED', 'PAID']);
  var withdrawRows = filterRows(withdrawsData, 7, 6, ['COMPLETED', 'PAID']);

  var sellMoney = 0; sellRows.forEach(function(r) { sellMoney += parseFloat(r[3]) || 0; });
  var tradeinMoney = 0; tradeinRows.forEach(function(r) { tradeinMoney += parseFloat(r[6]) || 0; });
  var exchangeMoney = 0; exchangeRows.forEach(function(r) { exchangeMoney += parseFloat(r[6]) || 0; });
  var switchMoney = 0; switchRows.forEach(function(r) { switchMoney += parseFloat(r[6]) || 0; });
  var freeExMoney = 0; freeExRows.forEach(function(r) { freeExMoney += parseFloat(r[5]) || 0; });

  var salesTotal = sellMoney + tradeinMoney + exchangeMoney + switchMoney + freeExMoney;
  var salesTotalTx = sellRows.length + tradeinRows.length + exchangeRows.length + switchRows.length + freeExRows.length;

  var salesOldGIn = 0, salesNewGOut = 0;
  sellRows.forEach(function(r) { salesNewGOut += calcG(r[2]); });
  tradeinRows.forEach(function(r) { salesOldGIn += calcG(r[2]); salesNewGOut += calcG(r[3]); });
  exchangeRows.forEach(function(r) { salesOldGIn += calcG(r[2]); salesNewGOut += calcG(r[3]); });
  switchRows.forEach(function(r) { salesOldGIn += calcG(r[2]); salesNewGOut += calcG(r[3]); });
  freeExRows.forEach(function(r) { salesOldGIn += calcG(r[2]); salesNewGOut += calcG(r[3]); });

  var bbMoney = 0; buybackRows.forEach(function(r) { bbMoney += parseFloat(r[6]) || parseFloat(r[3]) || 0; });
  var bbOldGIn = 0; buybackRows.forEach(function(r) { bbOldGIn += calcG(r[2]); });

  var wdMoney = 0; withdrawRows.forEach(function(r) { wdMoney += parseFloat(r[4]) || 0; });
  var wdNewGOut = 0; withdrawRows.forEach(function(r) { wdNewGOut += calcG(r[2]); });

  var totalOldGIn = salesOldGIn + bbOldGIn;
  var totalNewGOut = salesNewGOut + wdNewGOut;
  var netSellBaht = (totalNewGOut - totalOldGIn) / 15;

  var salesGoldBaht = (salesNewGOut - salesOldGIn) / 15;
  var salesTotalPerBaht = salesGoldBaht > 0 ? Math.round(salesTotal / salesGoldBaht) : 0;

  var bbGoldBaht = bbOldGIn / 15;
  var bbTotalPerBaht = bbGoldBaht > 0 ? Math.round(bbMoney / bbGoldBaht) : 0;

  var boxStyle = 'background:var(--bg-secondary);border:1px solid var(--border-color);border-radius:12px;padding:20px;';
  var netColor = netSellBaht >= 0 ? '#4caf50' : '#f44336';

  var html = '<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(250px,1fr));gap:15px;margin-bottom:25px;">';

  html += '<div style="' + boxStyle + '">';
  html += '<h3 style="color:var(--gold-primary);margin-bottom:8px;">⚖ NET SELL</h3>';
  html += '<p style="font-size:24px;margin:8px 0;font-weight:bold;color:' + netColor + ';">' + netSellBaht.toFixed(2) + ' <span style="font-size:13px;">บาท</span></p>';
  html += '<div style="border-top:1px solid var(--border-color);margin:6px 0;padding-top:6px;font-size:11px;color:var(--text-secondary);line-height:1.6;">';
  html += 'New Out ทั้งหมด: ' + totalNewGOut.toFixed(2) + ' g<br>';
  html += 'Old In ทั้งหมด: ' + totalOldGIn.toFixed(2) + ' g<br>';
  html += 'Net: ' + (totalNewGOut - totalOldGIn).toFixed(2) + ' g ÷ 15';
  html += '</div></div>';

  html += '<div style="' + boxStyle + '">';
  html += '<h3 style="color:var(--gold-primary);margin-bottom:8px;">💰 SALES</h3>';
  html += '<p style="font-size:18px;margin:3px 0;font-weight:bold;">Total: ' + formatNumber(Math.round(salesTotal)) + ' <span style="font-size:12px;">LAK</span></p>';
  html += '<p style="font-size:13px;margin:3px 0;">GOLD Amount: <b>' + salesGoldBaht.toFixed(2) + '</b> <span style="font-size:11px;">บาท</span></p>';
  html += '<p style="font-size:13px;margin:3px 0;">Total/Amount: <b>' + formatNumber(salesTotalPerBaht) + '</b> <span style="font-size:11px;">LAK/บาท</span></p>';
  html += '<p style="font-size:11px;color:var(--text-secondary);margin:2px 0;">Tx: <b>' + salesTotalTx + '</b></p>';
  html += '<div style="border-top:1px solid var(--border-color);margin:6px 0;padding-top:6px;font-size:11px;color:var(--text-secondary);line-height:1.6;">';
  html += 'Sell: ' + formatNumber(Math.round(sellMoney)) + ' (' + sellRows.length + ')<br>';
  html += 'Trade-in: ' + formatNumber(Math.round(tradeinMoney)) + ' (' + tradeinRows.length + ')<br>';
  html += 'Exchange: ' + formatNumber(Math.round(exchangeMoney)) + ' (' + exchangeRows.length + ')';
  if (switchRows.length > 0) html += '<br>Switch: ' + formatNumber(Math.round(switchMoney)) + ' (' + switchRows.length + ')';
  if (freeExRows.length > 0) html += '<br>Free-Ex: ' + formatNumber(Math.round(freeExMoney)) + ' (' + freeExRows.length + ')';
  html += '</div>';
  html += '<div style="border-top:1px solid var(--border-color);margin:6px 0;padding-top:6px;font-size:12px;">';
  html += '<span style="color:#ff9800;">◀ Old In: ' + salesOldGIn.toFixed(2) + ' g</span><br>';
  html += '<span style="color:#4caf50;">▶ New Out: ' + salesNewGOut.toFixed(2) + ' g</span>';
  html += '</div></div>';

  html += '<div style="' + boxStyle + '">';
  html += '<h3 style="color:var(--gold-primary);margin-bottom:8px;">🔄 BUYBACK</h3>';
  html += '<p style="font-size:18px;margin:3px 0;font-weight:bold;">Total: ' + formatNumber(Math.round(bbMoney)) + ' <span style="font-size:12px;">LAK</span></p>';
  html += '<p style="font-size:13px;margin:3px 0;">GOLD Amount: <b>' + bbGoldBaht.toFixed(2) + '</b> <span style="font-size:11px;">บาท</span></p>';
  html += '<p style="font-size:13px;margin:3px 0;">Total/Amount: <b>' + formatNumber(bbTotalPerBaht) + '</b> <span style="font-size:11px;">LAK/บาท</span></p>';
  html += '<p style="font-size:11px;color:var(--text-secondary);margin:2px 0;">Tx: <b>' + buybackRows.length + '</b></p>';
  html += '<div style="border-top:1px solid var(--border-color);margin:6px 0;padding-top:6px;font-size:12px;">';
  html += '<span style="color:#ff9800;">◀ Old In: ' + bbOldGIn.toFixed(2) + ' g</span>';
  html += '</div></div>';

  html += '<div style="' + boxStyle + '">';
  html += '<h3 style="color:var(--gold-primary);margin-bottom:8px;">📤 WITHDRAW</h3>';
  html += '<p style="font-size:18px;margin:3px 0;font-weight:bold;">' + formatNumber(Math.round(wdMoney)) + ' <span style="font-size:12px;">LAK</span></p>';
  html += '<p style="font-size:11px;color:var(--text-secondary);margin:2px 0;">Tx: <b>' + withdrawRows.length + '</b></p>';
  html += '<div style="border-top:1px solid var(--border-color);margin:6px 0;padding-top:6px;font-size:12px;">';
  html += '<span style="color:#4caf50;">▶ New Out: ' + wdNewGOut.toFixed(2) + ' g</span>';
  html += '</div></div>';

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
  oldInG += sumOldIn(buybacksData, 2, 10, 9, ['COMPLETED']);

  var boxStyle = 'background:var(--bg-secondary);border:1px solid var(--border-color);border-radius:12px;padding:20px;text-align:center;';
  var html = '<div style="display:grid;grid-template-columns:1fr 1fr;gap:15px;margin-bottom:20px;">';
  html += '<div style="' + boxStyle + '"><div style="color:var(--text-secondary);font-size:12px;margin-bottom:8px;">NEW OUT (ทองใหม่ออก)</div><div style="font-size:24px;font-weight:700;color:#f44336;">' + newOutG.toFixed(2) + ' g</div></div>';
  html += '<div style="' + boxStyle + '"><div style="color:var(--text-secondary);font-size:12px;margin-bottom:8px;">OLD IN (ทองเก่ารับเข้า)</div><div style="font-size:24px;font-weight:700;color:#4caf50;">' + oldInG.toFixed(2) + ' g</div></div>';
  html += '</div>';
  container.innerHTML = html;
}

function renderLRGoldTable(stockMoveNewData, stockMoveOldData, dateFrom, dateTo) {
  var container = document.getElementById('lrGoldTable');
  if (!container) return;
  var products = ['G01', 'G02', 'G03', 'G04', 'G05', 'G06', 'G07'];
  var names = { 'G01': '10 บาท', 'G02': '5 บาท', 'G03': '2 บาท', 'G04': '1 บาท', 'G05': '2 สลึง', 'G06': '1 สลึง', 'G07': '1 กรัม' };
  var newOut = {}, newIn = {}, oldOut = {}, oldIn = {};
  products.forEach(function(p) { newOut[p] = 0; newIn[p] = 0; oldOut[p] = 0; oldIn[p] = 0; });

  function parseMove(data, outMap, inMap) {
    if (!data || data.length <= 1) return;
    for (var i = 1; i < data.length; i++) {
      if (!lrInRange(data[i][0], dateFrom, dateTo)) continue;
      var dir = String(data[i][5] || '').trim();
      try {
        var items = typeof data[i][3] === 'string' ? JSON.parse(data[i][3]) : data[i][3];
        if (!Array.isArray(items)) continue;
        items.forEach(function(x) {
          if (dir === 'OUT' && outMap[x.productId] !== undefined) outMap[x.productId] += (x.qty || 0);
          if (dir === 'IN' && inMap[x.productId] !== undefined) inMap[x.productId] += (x.qty || 0);
        });
      } catch(e) {}
    }
  }

  parseMove(stockMoveNewData, newOut, newIn);
  parseMove(stockMoveOldData, oldOut, oldIn);

  var thStyle = 'background:#2d2d2d;color:#d4af37;border:1px solid rgba(212,175,55,0.5);padding:10px 8px;font-size:12px;text-align:center;font-weight:700;';
  var tdStyle = 'border:1px solid var(--border-color);padding:8px;text-align:center;font-size:13px;';

  var html = '<div style="background:var(--bg-secondary);border:1px solid var(--border-color);border-radius:12px;padding:16px;">';
  html += '<h3 style="color:var(--gold-primary);font-size:16px;margin-bottom:12px;">รายละเอียดทองแต่ละ Product</h3>';
  html += '<div class="table-container"><table><thead><tr>';
  html += '<th style="' + thStyle + '">Product</th>';
  html += '<th style="' + thStyle + '">New Out</th>';
  html += '<th style="' + thStyle + '">New In</th>';
  html += '<th style="' + thStyle + '">Old Out</th>';
  html += '<th style="' + thStyle + '">Old In</th>';
  html += '</tr></thead><tbody>';

  var tNewOut = 0, tNewIn = 0, tOldOut = 0, tOldIn = 0;
  products.forEach(function(p) {
    tNewOut += newOut[p]; tNewIn += newIn[p]; tOldOut += oldOut[p]; tOldIn += oldIn[p];
    html += '<tr>';
    html += '<td style="' + tdStyle + 'font-weight:600;">' + names[p] + ' (' + p + ')</td>';
    html += '<td style="' + tdStyle + 'color:' + (newOut[p] > 0 ? '#f44336' : 'var(--text-secondary)') + ';">' + newOut[p] + '</td>';
    html += '<td style="' + tdStyle + 'color:' + (newIn[p] > 0 ? '#4caf50' : 'var(--text-secondary)') + ';">' + newIn[p] + '</td>';
    html += '<td style="' + tdStyle + 'color:' + (oldOut[p] > 0 ? '#f44336' : 'var(--text-secondary)') + ';">' + oldOut[p] + '</td>';
    html += '<td style="' + tdStyle + 'color:' + (oldIn[p] > 0 ? '#4caf50' : 'var(--text-secondary)') + ';">' + oldIn[p] + '</td>';
    html += '</tr>';
  });

  html += '<tr style="background:rgba(212,175,55,0.1);font-weight:700;">';
  html += '<td style="' + tdStyle + 'color:var(--gold-primary);font-weight:700;">รวม</td>';
  html += '<td style="' + tdStyle + 'color:#f44336;font-weight:700;">' + tNewOut + '</td>';
  html += '<td style="' + tdStyle + 'color:#4caf50;font-weight:700;">' + tNewIn + '</td>';
  html += '<td style="' + tdStyle + 'color:#f44336;font-weight:700;">' + tOldOut + '</td>';
  html += '<td style="' + tdStyle + 'color:#4caf50;font-weight:700;">' + tOldIn + '</td>';
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
    var spinner = document.getElementById('salesInfoSpinner');
    var content = document.getElementById('salesInfoContent');
    if (spinner) spinner.style.display = 'block';
    if (content) content.style.display = 'none';

    var sheetName = currentUser.nickname;
    var batchRanges = [sheetName + '!A:I', sheetName + '_Gold!A:F'];
    var userData = [], goldData = [];
    try {
      var br = await callAppsScript('BATCH_READ', { ranges: JSON.stringify(batchRanges) });
      if (br && br.success && br.data) {
        userData = br.data[sheetName + '!A:I'] || [];
        goldData = br.data[sheetName + '_Gold!A:F'] || [];
      }
    } catch(e) {}
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

    document.getElementById('siCashLAK').textContent = formatNumber(cashLAK);
    document.getElementById('siCashTHB').textContent = formatNumber(cashTHB);
    document.getElementById('siCashUSD').textContent = formatNumber(cashUSD);
    document.getElementById('siSellPrice').textContent = formatNumber(sellPrice);
    document.getElementById('siBuybackPrice').textContent = formatNumber(buybackPrice);
    document.getElementById('siThbSell').textContent = formatNumber(currentExchangeRates.THB_Sell || 0);
    document.getElementById('siUsdSell').textContent = formatNumber(currentExchangeRates.USD_Sell || 0);
    document.getElementById('siThbBuy').textContent = formatNumber(currentExchangeRates.THB_Buy || 0);
    document.getElementById('siUsdBuy').textContent = formatNumber(currentExchangeRates.USD_Buy || 0);

    var products = ['G01','G02','G03','G04','G05','G06','G07'];
    var pNames = {'G01':'10 บาท','G02':'5 บาท','G03':'2 บาท','G04':'1 บาท','G05':'2 สลึง','G06':'1 สลึง','G07':'1 กรัม'};
    var goldQty = {};
    products.forEach(function(p) { goldQty[p] = 0; });
    if (goldData && goldData.length > 1) {
      for (var gi2 = 1; gi2 < goldData.length; gi2++) {
        var pid2 = String(goldData[gi2][0] || '').trim();
        var qty2 = parseFloat(goldData[gi2][1]) || 0;
        if (goldQty[pid2] !== undefined) goldQty[pid2] += qty2;
      }
    }
    var tblHtml = '<table style="width:100%;font-size:12px;border-collapse:collapse;">';
    tblHtml += '<tr><th style="text-align:left;padding:3px 0;color:var(--text-secondary);font-weight:600;">PRODUCT</th><th style="text-align:center;padding:3px 0;color:var(--text-secondary);font-weight:600;">UNIT</th></tr>';
    products.forEach(function(p) {
      var q = goldQty[p];
      var c = q > 0 ? 'color:var(--gold-primary);font-weight:600;' : 'color:var(--text-secondary);';
      tblHtml += '<tr style="border-top:1px solid var(--border-color);"><td style="padding:6px 0;">' + pNames[p] + '</td><td style="text-align:center;padding:6px 0;' + c + '">' + q + '</td></tr>';
    });
    tblHtml += '</table>';
    document.getElementById('siOldGoldTable').innerHTML = tblHtml;

    if (spinner) spinner.style.display = 'none';
    if (content) content.style.display = 'block';
  } catch(e) {
    console.error('loadSalesInfoBar error:', e);
  }
}