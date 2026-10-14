/**
 * IBKR – map Import rows onto Master (stocks + cash only).
 * Parses DateTime YYYYMMDD;HHMMSS into Trade Date, Trade Time, Trade Time Stamp.
 * Skips USD.CAD. Account=LT, Broker=IBKR.
 */
function mapIBKRImportToMaster() {
  const ss = SpreadsheetApp.getActive();
  const importSheet = ss.getSheetByName('Import');
  const masterSheet = ss.getSheetByName('Master');
  if (!importSheet || !masterSheet) throw new Error('Import or Master sheet missing');

  const importValues = importSheet.getDataRange().getValues();
  if (importValues.length < 2) {
    SpreadsheetApp.getUi().alert('No data on Import');
    return;
  }

  const headers = importValues[0].map(h => String(h).trim());
  const col = name => headers.indexOf(name);

  const idx = {
    symbol: col('Symbol'),
    dateTime: col('DateTime'),
    tradeDate: col('TradeDate'),
    txnType: col('TransactionType'),
    quantity: col('Quantity'),
    tradePrice: col('TradePrice'),
    proceeds: col('Proceeds'),
    netCash: col('NetCash'),
    commission: col('IBCommission'),
    notes: col('Notes/Codes'),
    assetClass: col('AssetClass'),
    description: col('Description')
  };

  const masterHeaders = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getValues()[0];
  const mCol = name => masterHeaders.indexOf(name);

  const m = {
    account: mCol('Account'),
    ticker: mCol('Ticker'),
    tradeDate: mCol('Trade Date'),
    tradeTime: mCol('Trade Time'),
    tradeTimeStamp: mCol('Trade Time Stamp'),
    action: mCol('Action'),
    quantity: mCol('Quantity'),
    entryPrice: mCol('Entry Price'),
    totalCost: mCol('Total Cost'),
    fees: mCol('Fees & Comm'),
    notes: mCol('Notes'),
    transferType: mCol('Transfer Type'),
    cashFlow: mCol('Cash Flow Direction'),
    broker: mCol('Broker')
  };

  // Parse IBKR DateTime "20220329;145615" → {date, time, stamp}
  function parseIBKRDateTime(raw) {
    const s = String(raw || '').trim();
    if (!s || s.indexOf(';') === -1) return { date: '', time: '', stamp: '' };
    const parts = s.split(';');
    const d = parts[0]; // YYYYMMDD
    const t = parts[1]; // HHMMSS
    if (d.length !== 8 || t.length < 4) return { date: '', time: '', stamp: '' };

    const yyyy = d.substring(0, 4);
    const mm = d.substring(4, 6);
    const dd = d.substring(6, 8);
    const hh = t.substring(0, 2);
    const min = t.substring(2, 4);
    const ss = t.length >= 6 ? t.substring(4, 6) : '00';

    const dateStr = yyyy + '-' + mm + '-' + dd;
    const timeStr = hh + ':' + min + ':' + ss;
    const stamp = new Date(dateStr + 'T' + timeStr);
    return { date: dateStr, time: timeStr, stamp: isNaN(stamp.getTime()) ? '' : stamp };
  }

  const out = [];

  for (let r = 1; r < importValues.length; r++) {
    const row = importValues[r];
    const symbol = String(row[idx.symbol] || '').trim();
    const asset = String(row[idx.assetClass] || '').trim();
    const desc = String(row[idx.description] || '').trim();
    const notes = String(row[idx.notes] || '').trim();

    if (symbol === 'USD.CAD') continue;

    const masterRow = new Array(masterHeaders.length).fill('');
    masterRow[m.account] = 'LT';
    masterRow[m.broker] = 'IBKR';

    if (asset === 'STK' && symbol && symbol !== 'USD.CAD') {
      // Stock trade
      const parsed = parseIBKRDateTime(row[idx.dateTime]);
      masterRow[m.ticker] = symbol;
      masterRow[m.tradeDate] = parsed.date;
      masterRow[m.tradeTime] = parsed.time;
      masterRow[m.tradeTimeStamp] = parsed.stamp;
      masterRow[m.action] = String(row[idx.txnType] || '').trim();
      masterRow[m.quantity] = row[idx.quantity];
      masterRow[m.entryPrice] = row[idx.tradePrice];
      masterRow[m.totalCost] = row[idx.proceeds] || row[idx.netCash];
      masterRow[m.fees] = row[idx.commission];
      masterRow[m.notes] = notes;

    } else if (desc.indexOf('CASH RECEIPTS') !== -1 || notes.indexOf('CASH RECEIPTS') !== -1) {
      // Cash – amount is in TradeDate column; try to find a real date elsewhere
      masterRow[m.transferType] = 'Deposits/Withdrawals';
      masterRow[m.cashFlow] = 'In';
      masterRow[m.totalCost] = row[idx.tradeDate];
      masterRow[m.notes] = desc || notes;

      // Look for a YYYYMMDD date in DateTime or other fields; fall back to blank
      let cashDate = '';
      const dtRaw = String(row[idx.dateTime] || '').trim();
      if (/^\d{8}/.test(dtRaw)) {
        cashDate = dtRaw.substring(0, 4) + '-' + dtRaw.substring(4, 6) + '-' + dtRaw.substring(6, 8);
      }
      masterRow[m.tradeDate] = cashDate;
      masterRow[m.tradeTimeStamp] = cashDate ? new Date(cashDate) : '';

    } else {
      continue;
    }

    out.push(masterRow);
  }

  if (out.length === 0) {
    SpreadsheetApp.getUi().alert('No rows mapped');
    return;
  }

  const masterLast = masterSheet.getLastRow();
  if (masterLast > 1) {
    masterSheet.getRange(2, 1, masterLast - 1, masterSheet.getLastColumn()).clearContent();
  }
  masterSheet.getRange(2, 1, out.length, out[0].length).setValues(out);

  SpreadsheetApp.getUi().alert('Mapped ' + out.length + ' rows to Master (LT / IBKR)');
}