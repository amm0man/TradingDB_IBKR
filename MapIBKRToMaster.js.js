/**
 * IBKR – map Import rows onto Master (stocks + cash only).
 * Parses DateTime YYYYMMDD;HHMMSS into Trade Date, Trade Time, Trade Time Stamp.
 * Skips USD.CAD. Account=LT, Broker=IBKR.
 */
function mapIBKRImportToMaster() {
  const ss = SpreadsheetApp.getActive();
  const importSheet = ss.getSheetByName("Import");
  const masterSheet = ss.getSheetByName("Master");
  if (!importSheet || !masterSheet)
    throw new Error("Import or Master sheet missing");

  const importValues = importSheet.getDataRange().getValues();
  if (importValues.length < 2) {
    SpreadsheetApp.getUi().alert("No data on Import");
    return;
  }

  const headers = importValues[0].map((h) => String(h).trim());
  const col = (name) => headers.indexOf(name);

  const idx = {
    symbol: col("Symbol"),
    dateTime: col("DateTime"),
    tradeDate: col("TradeDate"),
    txnType: col("TransactionType"),
    quantity: col("Quantity"),
    tradePrice: col("TradePrice"),
    proceeds: col("Proceeds"),
    netCash: col("NetCash"),
    commission: col("IBCommission"),
    notes: col("Notes/Codes"),
    assetClass: col("AssetClass"),
    description: col("Description"),
    exchange: col("Exchange"),
  };

  const masterHeaders = masterSheet
    .getRange(1, 1, 1, masterSheet.getLastColumn())
    .getValues()[0];
  const mCol = (name) => masterHeaders.indexOf(name);

  const m = {
    account: mCol("Account"),
    ticker: mCol("Ticker"),
    tradeDate: mCol("Trade Date"),
    tradeTime: mCol("Trade Time"),
    tradeTimeStamp: mCol("Trade Time Stamp"),
    action: mCol("Action"),
    quantity: mCol("Quantity"),
    entryPrice: mCol("Entry Price"),
    totalCost: mCol("Total Cost"),
    fees: mCol("Fees & Comm"),
    notes: mCol("Notes"),
    transferType: mCol("Transfer Type"),
    cashFlow: mCol("Cash Flow Direction"),
    broker: mCol("Broker"),
    strategyType: mCol("Strategy Type"),
    tradeType: mCol("Trade Type"),
    corpActions: mCol("Corporate Actions"),
  };

  // Parse IBKR DateTime "20220329;145615" → {date, time, stamp}
  function parseIBKRDateTime(raw) {
    const s = String(raw || "").trim();
    if (!s || s.indexOf(";") === -1) return { date: "", time: "", stamp: "" };
    const parts = s.split(";");
    const d = parts[0]; // YYYYMMDD
    const t = parts[1]; // HHMMSS
    if (d.length !== 8 || t.length < 4)
      return { date: "", time: "", stamp: "" };

    const yyyy = d.substring(0, 4);
    const mm = d.substring(4, 6);
    const dd = d.substring(6, 8);
    const hh = t.substring(0, 2);
    const min = t.substring(2, 4);

    // Build a real Date object (seconds forced to 0 so it matches TOS style)
    const stamp = new Date(
      Number(yyyy),
      Number(mm) - 1,
      Number(dd),
      Number(hh),
      Number(min),
      0,
    );

    return {
      date: stamp, // real date – Sheets/Data Studio will treat it as a date
      time: stamp, // same value; we will format the column to show only time
      stamp: stamp, // full date+time
    };
  }
  const out = [];

  for (let r = 1; r < importValues.length; r++) {
    const row = importValues[r];
    const symbol = String(row[idx.symbol] || "").trim();
    const asset = String(row[idx.assetClass] || "").trim();
    const desc = String(row[idx.description] || "").trim();
    const notes = String(row[idx.notes] || "").trim();

    if (symbol === "USD.CAD") continue;

    const masterRow = new Array(masterHeaders.length).fill("");
    masterRow[m.account] = "LT";
    masterRow[m.broker] = "IBKR";

    const txnType = String(row[idx.txnType] || "").trim();
    const isDividend =
      txnType === "Dividends" || desc.indexOf("CASH DIVIDEND") !== -1;
    const isCashReceipt =
      desc.indexOf("CASH RECEIPTS") !== -1 ||
      notes.indexOf("CASH RECEIPTS") !== -1;

    if (isDividend) {
      // Dividend – amount is under TradeDate (same shift pattern as cash receipts)
      masterRow[m.account] = "LT";
      masterRow[m.broker] = "IBKR";
      masterRow[m.ticker] = String(symbol || "").toUpperCase();
      masterRow[m.transferType] = "DIVIDEND";
      masterRow[m.cashFlow] = "IN";
      masterRow[m.totalCost] = row[idx.tradeDate]; // the real cash amount
      masterRow[m.notes] = String(desc || notes || "").toUpperCase();
      masterRow[m.strategyType] = "LONG STOCK";
      masterRow[m.tradeType] = "STOCK";

     
      // Try to pull a clean date from DateTime
      let divDate = "";
      const dtRaw = String(row[idx.dateTime] || "").trim();
      if (/^\d{8}/.test(dtRaw)) {
        divDate =
          dtRaw.substring(0, 4) +
          "-" +
          dtRaw.substring(4, 6) +
          "-" +
          dtRaw.substring(6, 8);
      }
      if (divDate) {
        const parts = divDate.split("-");
        const divStamp = new Date(
          Number(parts[0]),
          Number(parts[1]) - 1,
          Number(parts[2]),
        );
        masterRow[m.tradeDate] = divStamp;
        masterRow[m.tradeTimeStamp] = divStamp;
      }
    } else if (asset === "STK" && symbol && symbol !== "USD.CAD") {
      // Stock trade (includes DRIP share purchases)
      const parsed = parseIBKRDateTime(row[idx.dateTime]);
      const qty = Number(row[idx.quantity]) || 0;
      const exchange = String(row[idx.exchange] || "")
        .trim()
        .toUpperCase();
      const isDrip = exchange.indexOf("IBDRIP") !== -1;

      masterRow[m.ticker] = String(symbol).toUpperCase();
      masterRow[m.tradeDate] = parsed.date;
      masterRow[m.tradeTime] = parsed.time;
      masterRow[m.tradeTimeStamp] = parsed.stamp;

      if (qty > 0) {
        masterRow[m.action] = "BUY TO OPEN";
      } else if (qty < 0) {
        masterRow[m.action] = "SELL TO CLOSE";
      } else {
        masterRow[m.action] = "EXCHTRADE";
      }

      masterRow[m.quantity] = qty;
      masterRow[m.entryPrice] = row[idx.tradePrice];
      masterRow[m.totalCost] = row[idx.proceeds] || row[idx.netCash];
      masterRow[m.fees] = row[idx.commission];
      masterRow[m.notes] = String(notes || "").toUpperCase();

      // Standard fields for querying (same as TOS)
      masterRow[mCol("Strategy Type")] = "LONG STOCK";
      masterRow[mCol("Trade Type")] = "STOCK";

      // Flag DRIP
      if (isDrip) {
        masterRow[mCol("Corporate Actions")] = "DRIP";
      }
       if (isDrip) {
        masterRow[m.corpActions] = "DRIP";
      }

    } else if (isCashReceipt) {
      // existing cash-receipt block (unchanged)
      masterRow[m.transferType] = "DEPOSITS/WITHDRAWALS";
      masterRow[m.cashFlow] = "IN";
      masterRow[m.totalCost] = row[idx.tradeDate];
      masterRow[m.notes] = String(desc || notes || "").toUpperCase();

      let cashDate = "";
      const dtRaw = String(row[idx.dateTime] || "").trim();
      if (/^\d{8}/.test(dtRaw)) {
        cashDate =
          dtRaw.substring(0, 4) +
          "-" +
          dtRaw.substring(4, 6) +
          "-" +
          dtRaw.substring(6, 8);
      }
      if (cashDate) {
        const parts = cashDate.split("-");
        const cashStamp = new Date(
          Number(parts[0]),
          Number(parts[1]) - 1,
          Number(parts[2]),
        );
        masterRow[m.tradeDate] = cashStamp;
        masterRow[m.tradeTimeStamp] = cashStamp;
      }
    } else {
      continue;
    }

    // Look for a YYYYMMDD date in DateTime or other fields; fall back to blank
    let cashDate = "";
    const dtRaw = String(row[idx.dateTime] || "").trim();
    if (/^\d{8}/.test(dtRaw)) {
      cashDate =
        dtRaw.substring(0, 4) +
        "-" +
        dtRaw.substring(4, 6) +
        "-" +
        dtRaw.substring(6, 8);
    }

    if (cashDate) {
      const parts = cashDate.split("-");
      const cashStamp = new Date(
        Number(parts[0]),
        Number(parts[1]) - 1,
        Number(parts[2]),
      );
      masterRow[m.tradeDate] = cashStamp;
      masterRow[m.tradeTimeStamp] = cashStamp;
    }

    out.push(masterRow);
  }

  if (out.length === 0) {
    SpreadsheetApp.getUi().alert("No rows mapped");
    return;
  }

  const masterLast = masterSheet.getLastRow();
  if (masterLast > 1) {
    masterSheet
      .getRange(2, 1, masterLast - 1, masterSheet.getLastColumn())
      .clearContent();
  }
  masterSheet.getRange(2, 1, out.length, out[0].length).setValues(out);

  SpreadsheetApp.getUi().alert(
    "Mapped " + out.length + " rows to Master (LT / IBKR)",
  );
}
