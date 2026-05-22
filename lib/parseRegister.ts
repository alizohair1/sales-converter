// lib/parseRegister.ts
// Parses raw Sales Register (.xls/.xlsx) into 10-minute interval row objects
//
// SR structure per order header row (columns):
//   col0 = "DATE :"
//   col1 = date (as date string/timestamp)
//   col2 = shift letter (A or B)
//   col3 = HOUR (0-23)
//   col4 = " : "
//   col5 = MINUTE (0-59)
//   col7 = till type (TAKE AWAY / HOME DEL)
//   col9 = customer name
//   col11 = payment method
//   col12 = receipt number
//
// Item row columns:
//   col0 = item code
//   col2 = item name
//   col5 = qty
//   col7 = rate (unit price)
//   col8 = amount (qty * rate)

export interface SaleRow {
  date: string;
  time: string;
  till: string;
  customer: string;
  payment: string;
  items: Record<string, number>;        // qty mode
  amounts: Record<string, number>;      // amount mode
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

export function parseSalesRegister(rawRows: (string | number | null)[][]): OrderRecord[] {
  const orders: OrderRecord[] = [];
  let current: OrderRecord | null = null;

  for (let i = 0; i < rawRows.length; i++) {
    const row = rawRows[i];
    const cell0 = String(row[0] ?? '').trim();

    // ── Order header ─────────────────────────────────────────────────────────
    if (cell0 === 'DATE :' || cell0.startsWith('DATE :')) {
      if (current) orders.push(current);

      const rawDate = row[1];
      const dateStr = formatDate(rawDate);

      const hour         = safeInt(row[3], 0);
      const minute       = safeInt(row[5], 0);
      const minuteBucket = Math.floor(minute / 10) * 10;
      const time         = `${pad(hour)}:${pad(minute)}`;
      const bucket       = `${pad(hour)}:${pad(minuteBucket)}`;

      const till     = String(row[7]  ?? '').trim();
      const customer = String(row[9]  ?? '').trim();
      const payment  = String(row[11] ?? '').trim();
      const receipt  = String(row[12] ?? '').trim();

      current = { date: dateStr, time, timeBucket: bucket, till, customer, payment, receipt, items: [] };
      continue;
    }

    // ── Item row ─────────────────────────────────────────────────────────────
    if (current) {
      const code   = String(row[0] ?? '').trim();
      const name   = String(row[2] ?? '').trim();
      const qty    = Math.abs(safeFloat(row[5], 0)) || 1;
      const rate   = Math.abs(safeFloat(row[7], 0));
      const amount = Math.abs(safeFloat(row[8], 0));

      const isValidItem =
        code &&
        name &&
        name !== 'Item Description' &&
        name !== 'nan' &&
        name !== 'NaN' &&
        !cell0.startsWith('DATE :') &&
        String(row[4] ?? '').trim() !== 'G.AMOUNT :';

      if (isValidItem) {
        current.items.push({
          code,
          name: name.toUpperCase().replace(/\s+/g, ' '),
          qty,
          rate,
          amount,
        });
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
        amounts: {},
      });
    }

    const bucket = buckets.get(key)!;
    for (const item of order.items) {
      bucket.items[item.name]   = (bucket.items[item.name]   || 0) + item.qty;
      bucket.amounts[item.name] = (bucket.amounts[item.name] || 0) + item.amount;
    }
  }

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

export function buildOutputTable(
  rows: SaleRow[],
  allItems: string[],
  mode: 'qty' | 'amount' = 'qty'
): (string | number)[][] {
  const header = ['Date', 'Time\n(10-min)', 'Till', 'Customer', 'Payment', ...allItems];
  const data = rows.map(r => [
    r.date, r.time, r.till, r.customer, r.payment,
    ...allItems.map(item =>
      mode === 'amount'
        ? (r.amounts[item] || '')
        : (r.items[item]   || '')
    ),
  ]);
  return [header, ...data];
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function formatDate(raw: string | number | null | undefined): string {
  if (!raw) return '';
  const s = String(raw).trim();
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) return s;
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[3]}/${m[2]}/${m[1]}`;
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
