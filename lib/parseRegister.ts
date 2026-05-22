// lib/parseRegister.ts
// Parses raw Sales Register (.xls/.xlsx) into 10-minute interval row objects
// 
// SR structure per order header row (columns):
//   col0 = "DATE :"
//   col1 = date (as date string/timestamp)
//   col2 = shift letter (A or B)
//   col3 = HOUR (0-23) — this is the actual clock hour
//   col4 = " : "
//   col5 = MINUTE (0-59) — actual clock minute within the hour
//   col7 = till type (TAKE AWAY / HOME DEL)
//   col9 = customer name
//   col11 = payment method
//   col12 = receipt number

export interface SaleRow {
  date: string;
  time: string;
  till: string;
  customer: string;
  payment: string;
  items: Record<string, number>;
}

export interface OrderRecord {
  date: string;
  time: string;       // HH:MM actual time
  timeBucket: string; // HH:M0 floored to 10-min
  till: string;
  customer: string;
  payment: string;
  receipt: string;
  items: { code: string; name: string; qty: number }[];
}

export function parseSalesRegister(rawRows: (string | number | null)[][]): OrderRecord[] {
  const orders: OrderRecord[] = [];
  let current: OrderRecord | null = null;

  for (let i = 0; i < rawRows.length; i++) {
    const row = rawRows[i];
    const cell0 = String(row[0] ?? '').trim();

    // ── Order header ──────────────────────────────────────────────────────────
    if (cell0 === 'DATE :' || cell0.startsWith('DATE :')) {
      if (current) orders.push(current);

      // Date: col1 is a formatted date string like "2026-04-06" or a JS Date
      const rawDate = row[1];
      const dateStr = formatDate(rawDate);

      // Hour = col3, Minute = col5
      const hour   = safeInt(row[3], 0);
      const minute = safeInt(row[5], 0);
      const minuteBucket = Math.floor(minute / 10) * 10;
      const time   = `${pad(hour)}:${pad(minute)}`;
      const bucket = `${pad(hour)}:${pad(minuteBucket)}`;

      const till     = String(row[7]  ?? '').trim();
      const customer = String(row[9]  ?? '').trim();
      const payment  = String(row[11] ?? '').trim();
      const receipt  = String(row[12] ?? '').trim();

      current = { date: dateStr, time, timeBucket: bucket, till, customer, payment, receipt, items: [] };
      continue;
    }

    // ── Item row ──────────────────────────────────────────────────────────────
    if (current) {
      const code = String(row[0] ?? '').trim();
      const name = String(row[2] ?? '').trim();
      const qty  = Math.abs(safeFloat(row[6], 0)) || 1;

      const isValidItem =
        code &&
        name &&
        name !== 'Item Description' &&
        name !== 'nan' &&
        name !== 'NaN' &&
        !cell0.startsWith('DATE :') &&
        row[4] !== 'G.AMOUNT :' &&
        String(row[4] ?? '').trim() !== 'G.AMOUNT :';

      if (isValidItem) {
        current.items.push({ code, name: name.toUpperCase().replace(/\s+/g, ' '), qty });
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
        date: order.date,
        time: order.timeBucket,
        till: order.till,
        customer: order.customer,
        payment: order.payment,
        items: {},
      });
    }

    const bucket = buckets.get(key)!;
    for (const item of order.items) {
      bucket.items[item.name] = (bucket.items[item.name] || 0) + item.qty;
    }
  }

  // Sort by date then time
  return Array.from(buckets.values()).sort((a, b) => {
    if (a.date !== b.date) return a.date.localeCompare(b.date);
    return a.time.localeCompare(b.time);
  });
}

export function collectAllItems(rows: SaleRow[]): string[] {
  const set = new Set<string>();
  for (const r of rows) for (const k of Object.keys(r.items)) set.add(k);
  return Array.from(set).sort();
}

export function buildOutputTable(rows: SaleRow[], allItems: string[]): (string | number)[][] {
  const header = ['Date', 'Time\n(10-min)', 'Till', 'Customer', 'Payment', ...allItems];
  const data = rows.map(r => [
    r.date, r.time, r.till, r.customer, r.payment,
    ...allItems.map(item => r.items[item] || ''),
  ]);
  return [header, ...data];
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function formatDate(raw: string | number | null | undefined): string {
  if (!raw) return '';
  const s = String(raw).trim();
  // Already DD/MM/YYYY
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) return s;
  // YYYY-MM-DD
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[3]}/${m[2]}/${m[1]}`;
  // Excel serial number (SheetJS returns these as strings when raw:false is off)
  return s;
}

function safeInt(val: unknown, fallback: number): number {
  const n = parseFloat(String(val ?? '').replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? fallback : Math.floor(n);
}

function safeFloat(val: unknown, fallback: number): number {
  const n = parseFloat(String(val ?? '').replace(/[^0-9.\-]/g, ''));
  return isNaN(n) ? fallback : n;
}

function pad(n: number): string {
  return String(n).padStart(2, '0');
}
