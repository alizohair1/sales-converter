// lib/parseRegister.ts
// Parses raw Sales Register — supports:
//   • .xls / .xlsx  (SheetJS array-of-arrays, multi-sheet)
//   • .csv          (flat format: every row = header prefix + order info + one item)
//
// XLS/XLSX column layout per order header row:
//   col0='DATE :' col1=date col2=shift col3=HOUR col4=' : ' col5=MINUTE
//   col7=till col9=customer col11=payment col12=receipt
//   Item rows: col0=code col2=name col5=qty col7=rate col8=amount
//
// CSV column layout (every row repeats full prefix):
//   col0='SALE REGISTER' col1=report-range col2='Code' ... col7='Amount'
//   Then EITHER:
//     col8='DATE :' col9=date col10=shift col11=hour col12=' : ' col13=minute
//     col14='TILL :' col15=till col16='CUSTOMER :' col17=customer
//     col18='PAYMENT :' col19=payment col20=receipt
//     col21=itemCode col22=itemName col23=qty col24=rQty col25=rate col26=amount
//   OR (continuation item row, no DATE):
//     col8=itemCode col9=itemName col10=qty col11=rQty col12=rate col13=amount

export interface SaleRow {
  date: string;
  time: string;
  till: string;
  customer: string;
  payment: string;
  items: Record<string, number>;    // qty per item
  amounts: Record<string, number>;  // PKR amount per item
}

export interface OrderRecord {
  date: string;
  time: string;
  timeBucket: string;
  till: string;
  customer: string;
  payment: string;
  receipt: string;
  items: { code: string; name: string; qty: number; rate: number; amount: number }[];
}

// ── XLS/XLSX parser (array-of-arrays from SheetJS) ───────────────────────────
export function parseSalesRegister(rawRows: (string | number | null)[][]): OrderRecord[] {
  const orders: OrderRecord[] = [];
  let current: OrderRecord | null = null;

  for (const row of rawRows) {
    const cell0 = String(row[0] ?? '').trim();

    if (cell0 === 'DATE :' || cell0.startsWith('DATE :')) {
      if (current) orders.push(current);

      const dateStr      = formatDate(row[1]);
      const hour         = safeInt(row[3], 0);
      const minute       = safeInt(row[5], 0);
      const minuteBucket = Math.floor(minute / 10) * 10;
      const time         = `${pad(hour)}:${pad(minute)}`;
      const bucket       = `${pad(hour)}:${pad(minuteBucket)}`;

      current = {
        date: dateStr,
        time,
        timeBucket: bucket,
        till:     String(row[7]  ?? '').trim(),
        customer: String(row[9]  ?? '').trim(),
        payment:  String(row[11] ?? '').trim(),
        receipt:  String(row[12] ?? '').trim(),
        items: [],
      };
      continue;
    }

    if (current) {
      const code   = String(row[0] ?? '').trim();
      const name   = String(row[2] ?? '').trim();
      const qty    = Math.abs(safeFloat(row[5], 0)) || 1;
      const rate   = Math.abs(safeFloat(row[7], 0));
      const amount = Math.abs(safeFloat(row[8], 0));

      const isValid =
        code && name &&
        name !== 'Item Description' &&
        name !== 'nan' && name !== 'NaN' &&
        !cell0.startsWith('DATE :') &&
        String(row[4] ?? '').trim() !== 'G.AMOUNT :';

      if (isValid) {
        current.items.push({ code, name: cleanName(name), qty, rate, amount });
      }
    }
  }

  if (current) orders.push(current);
  return orders;
}

// ── CSV parser (flat format from POS export) ──────────────────────────────────
export function parseSalesRegisterCSV(csvText: string): OrderRecord[] {
  const orders: OrderRecord[] = [];
  let current: OrderRecord | null = null;

  // Split lines, parse each as CSV
  const lines = csvText.split(/\r?\n/);

  for (const line of lines) {
    if (!line.trim()) continue;
    const row = parseCSVLine(line);

    // Detect row type by col8
    const col8 = String(row[8] ?? '').trim();

    if (col8 === 'DATE :') {
      // ── Order header row ──
      if (current) orders.push(current);

      const dateStr      = formatDate(row[9]);   // col9 = date DD/MM/YYYY
      const hour         = safeInt(row[11], 0);  // col11 = hour
      const minute       = safeInt(row[13], 0);  // col13 = minute
      const minuteBucket = Math.floor(minute / 10) * 10;

      current = {
        date:       dateStr,
        time:       `${pad(hour)}:${pad(minute)}`,
        timeBucket: `${pad(hour)}:${pad(minuteBucket)}`,
        till:       String(row[15] ?? '').trim(),  // col15
        customer:   String(row[17] ?? '').trim(),  // col17
        payment:    String(row[19] ?? '').trim(),  // col19
        receipt:    String(row[20] ?? '').trim(),  // col20
        items: [],
      };

      // This row also contains one item at col21-26
      const code   = String(row[21] ?? '').trim();
      const name   = String(row[22] ?? '').trim();
      const qty    = Math.abs(safeFloat(row[23], 0)) || 1;
      const rate   = Math.abs(safeFloat(row[25], 0));
      const amount = Math.abs(safeFloat(row[26], 0));
      if (code && name && name !== 'Item Description') {
        current.items.push({ code, name: cleanName(name), qty, rate, amount });
      }

    } else if (col8 && col8 !== 'G.AMOUNT :' && col8 !== 'SALE REGISTER' && col8 !== 'Code') {
      // ── Continuation item row ── col8=code col9=name col10=qty col12=rate col13=amount
      if (!current) continue;

      const code   = col8;
      const name   = String(row[9]  ?? '').trim();
      const qty    = Math.abs(safeFloat(row[10], 0)) || 1;
      const rate   = Math.abs(safeFloat(row[12], 0));
      const amount = Math.abs(safeFloat(row[13], 0));

      if (name && name !== 'Item Description' && name !== 'G.AMOUNT :') {
        current.items.push({ code, name: cleanName(name), qty, rate, amount });
      }
    }
  }

  if (current) orders.push(current);
  return orders;
}

// ── Group into 10-min buckets ─────────────────────────────────────────────────
export function ordersTo10MinSummary(orders: OrderRecord[]): SaleRow[] {
  const buckets = new Map<string, SaleRow>();

  for (const order of orders) {
    const key = `${order.date}||${order.timeBucket}||${order.till}||${order.customer}||${order.payment}`;

    if (!buckets.has(key)) {
      buckets.set(key, {
        date:     order.date,
        time:     order.timeBucket,
        till:     order.till,
        customer: order.customer,
        payment:  order.payment,
        items:    {},
        amounts:  {},
      });
    }

    const bucket = buckets.get(key)!;
    for (const item of order.items) {
      bucket.items[item.name]   = (bucket.items[item.name]   || 0) + item.qty;
      bucket.amounts[item.name] = (bucket.amounts[item.name] || 0) + item.amount;
    }
  }

  return Array.from(buckets.values()).sort((a, b) => {
    const da = parseDateSort(a.date), db = parseDateSort(b.date);
    if (da !== db) return da.localeCompare(db);
    return a.time.localeCompare(b.time);
  });
}

export function collectAllItems(rows: SaleRow[]): string[] {
  const set = new Set<string>();
  for (const r of rows) for (const k of Object.keys(r.items)) set.add(k);
  return Array.from(set).sort();
}

export function buildOutputTable(
  rows: SaleRow[],
  allItems: string[],
  mode: 'qty' | 'amount' = 'qty'
): (string | number)[][] {
  const header = ['Date', 'Time\n(10-min)', 'Till', 'Customer', 'Payment', ...allItems];
  const data = rows.map(r => [
    r.date, r.time, r.till, r.customer, r.payment,
    ...allItems.map(item =>
      mode === 'amount' ? (r.amounts[item] || '') : (r.items[item] || '')
    ),
  ]);
  return [header, ...data];
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function cleanName(name: string): string {
  return name.toUpperCase().replace(/\s+/g, ' ').trim();
}

function formatDate(raw: string | number | null | undefined): string {
  if (!raw) return '';
  const s = String(raw).trim();
  // Already DD/MM/YYYY
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) return s;
  // YYYY-MM-DD
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[3]}/${m[2]}/${m[1]}`;
  return s;
}

// Convert DD/MM/YYYY → YYYY/MM/DD for sorting
function parseDateSort(d: string): string {
  const m = d.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m) return `${m[3]}/${m[2]}/${m[1]}`;
  return d;
}

function safeInt(val: unknown, fallback: number): number {
  const n = parseFloat(String(val ?? '').replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? fallback : Math.floor(n);
}

function safeFloat(val: unknown, fallback: number): number {
  const n = parseFloat(String(val ?? '').replace(/[^0-9,.\-]/g, '').replace(/,/g, ''));
  return isNaN(n) ? fallback : n;
}

function pad(n: number): string {
  return String(n).padStart(2, '0');
}

// Minimal CSV line parser (handles quoted fields with commas inside)
function parseCSVLine(line: string): string[] {
  const result: string[] = [];
  let current = '';
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      if (inQuotes && line[i + 1] === '"') { current += '"'; i++; }
      else inQuotes = !inQuotes;
    } else if (ch === ',' && !inQuotes) {
      result.push(current);
      current = '';
    } else {
      current += ch;
    }
  }
  result.push(current);
  return result;
}
