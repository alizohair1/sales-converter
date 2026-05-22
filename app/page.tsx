'use client';
import { useState, useCallback, useRef } from 'react';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import {
  parseSalesRegister,
  ordersTo10MinSummary,
  collectAllItems,
  buildOutputTable,
  OrderRecord,
  SaleRow,
} from '@/lib/parseRegister';

type Status = 'idle' | 'reading' | 'parsing' | 'done' | 'error';
type OutputMode = 'qty' | 'amount';

interface Stats {
  orders: number;
  rows: number;
  items: number;
  date: string;
}

export default function Home() {
  const [status, setStatus]       = useState<Status>('idle');
  const [stats, setStats]         = useState<Stats | null>(null);
  const [error, setError]         = useState('');
  const [preview, setPreview]     = useState<(string | number)[][]>([]);
  const [allItems, setAllItems]   = useState<string[]>([]);
  const [saleRows, setSaleRows]   = useState<SaleRow[]>([]);
  const [fileName, setFileName]   = useState('');
  const [isDragging, setIsDragging] = useState(false);
  const [outputMode, setOutputMode] = useState<OutputMode>('qty');
  const fileRef = useRef<HTMLInputElement>(null);

  const processFile = useCallback(async (file: File) => {
    setStatus('reading');
    setError('');
    setStats(null);
    setPreview([]);
    setFileName(file.name);

    try {
      const buffer = await file.arrayBuffer();
      setStatus('parsing');

      const workbook = XLSX.read(buffer, { type: 'array', cellDates: false });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];

      const rawData: (string | number | null)[][] = XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: null,
        raw: false,
      }) as (string | number | null)[][];

      const orders: OrderRecord[] = parseSalesRegister(rawData);
      const rows: SaleRow[]       = ordersTo10MinSummary(orders);
      const items                  = collectAllItems(rows);
      const table                  = buildOutputTable(rows, items, 'qty');

      const dateStr = orders[0]?.date || '';

      setStats({ orders: orders.length, rows: rows.length, items: items.length, date: dateStr });
      setAllItems(items);
      setSaleRows(rows);
      setPreview(table.slice(0, 8));
      setStatus('done');
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : 'Failed to parse file');
      setStatus('error');
    }
  }, []);

  // Rebuild preview when mode changes
  const handleModeChange = (mode: OutputMode) => {
    setOutputMode(mode);
    if (saleRows.length && allItems.length) {
      const table = buildOutputTable(saleRows, allItems, mode);
      setPreview(table.slice(0, 8));
    }
  };

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files[0];
    if (file) processFile(file);
  }, [processFile]);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) processFile(file);
  };

  const downloadXLSX = () => {
    if (!saleRows.length) return;
    const table = buildOutputTable(saleRows, allItems, outputMode);
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(table);

    ws['!cols'] = [
      { wch: 12 }, { wch: 10 }, { wch: 12 }, { wch: 28 }, { wch: 14 },
      ...allItems.map(() => ({ wch: 10 })),
    ];

    const sheetLabel = outputMode === 'amount' ? 'Sales by Amount' : 'Sales by 10-min Interval';
    XLSX.utils.book_append_sheet(wb, ws, sheetLabel);
    const buf = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const suffix = outputMode === 'amount' ? '_amount.xlsx' : '_10min.xlsx';
    const outName = fileName.replace(/\.(xls|xlsx)$/i, '') + suffix;
    saveAs(blob, outName);
  };

  const reset = () => {
    setStatus('idle');
    setStats(null);
    setError('');
    setPreview([]);
    setFileName('');
    setOutputMode('qty');
    if (fileRef.current) fileRef.current.value = '';
  };

  return (
    <div className="min-h-screen grid-bg relative overflow-hidden">
      {/* Ambient glow */}
      <div style={{
        position: 'fixed', top: '-20%', left: '50%', transform: 'translateX(-50%)',
        width: '600px', height: '400px',
        background: 'radial-gradient(ellipse, rgba(249,115,22,0.08) 0%, transparent 70%)',
        pointerEvents: 'none',
      }} />

      <div className="max-w-5xl mx-auto px-6 py-12">
        {/* Header */}
        <div className="mb-12 animate-fade-up">
          <div className="flex items-center gap-3 mb-4">
            <div style={{
              width: 40, height: 40, background: 'var(--accent)',
              clipPath: 'polygon(50% 0%, 100% 25%, 100% 75%, 50% 100%, 0% 75%, 0% 25%)',
            }} />
            <span className="mono text-xs tracking-widest" style={{ color: 'var(--muted)' }}>
              JT RESTAURANT SYSTEMS
            </span>
          </div>
          <h1 className="mono text-4xl font-bold mb-2" style={{ color: 'var(--text)' }}>
            Sales Register<br />
            <span style={{ color: 'var(--accent)' }}>Converter</span>
          </h1>
          <p style={{ color: 'var(--muted)', maxWidth: 480 }}>
            Upload your daily Sales Register (.xls or .xlsx) and instantly convert it
            to a 10-minute interval summary — ready to download.
          </p>
        </div>

        {/* Drop zone */}
        {(status === 'idle' || status === 'error') ? (
          <div
            className="animate-fade-up"
            style={{ animationDelay: '0.1s', opacity: 0 }}
            onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
            onDragLeave={() => setIsDragging(false)}
            onDrop={handleDrop}
            onClick={() => fileRef.current?.click()}
          >
            <div style={{
              border: `2px dashed ${isDragging ? 'var(--accent)' : 'var(--border)'}`,
              borderRadius: 16,
              padding: '60px 40px',
              textAlign: 'center',
              cursor: 'pointer',
              background: isDragging ? 'rgba(249,115,22,0.05)' : 'var(--surface)',
              transition: 'all 0.2s ease',
              boxShadow: isDragging ? '0 0 40px rgba(249,115,22,0.1)' : 'none',
            }}>
              <div style={{ fontSize: 48, marginBottom: 16 }}>📊</div>
              <p className="mono" style={{ color: 'var(--text)', fontSize: 18, marginBottom: 8 }}>
                Drop your Sales Register here
              </p>
              <p style={{ color: 'var(--muted)', fontSize: 14 }}>
                or click to browse · supports .xls and .xlsx
              </p>
              {error && (
                <div style={{
                  marginTop: 20, padding: '10px 16px',
                  background: 'rgba(239,68,68,0.1)', border: '1px solid rgba(239,68,68,0.3)',
                  borderRadius: 8, color: '#ef4444', fontSize: 14,
                }}>
                  ⚠ {error}
                </div>
              )}
            </div>
            <input
              ref={fileRef}
              type="file"
              accept=".xls,.xlsx"
              onChange={handleFileChange}
              style={{ display: 'none' }}
            />
          </div>
        ) : null}

        {/* Processing */}
        {(status === 'reading' || status === 'parsing') && (
          <div style={{
            background: 'var(--surface)',
            border: '1px solid var(--border)',
            borderRadius: 16,
            padding: '60px 40px',
            textAlign: 'center',
          }}>
            <div style={{
              width: 48, height: 48, border: '3px solid var(--border)',
              borderTopColor: 'var(--accent)', borderRadius: '50%',
              animation: 'spin-slow 0.8s linear infinite',
              margin: '0 auto 24px',
            }} />
            <p className="mono" style={{ color: 'var(--accent)', fontSize: 16 }}>
              {status === 'reading' ? 'Reading file...' : 'Parsing orders...'}
            </p>
            <p style={{ color: 'var(--muted)', fontSize: 13, marginTop: 8 }}>{fileName}</p>
          </div>
        )}

        {/* Done */}
        {status === 'done' && stats && (
          <div className="animate-fade-up">
            {/* Stats bar */}
            <div style={{
              display: 'grid', gridTemplateColumns: 'repeat(4, 1fr)',
              gap: 2, marginBottom: 24,
            }}>
              {[
                { label: 'DATE',         value: formatDisplayDate(stats.date) },
                { label: 'ORDERS',       value: stats.orders.toLocaleString() },
                { label: 'SUMMARY ROWS', value: stats.rows.toLocaleString() },
                { label: 'ITEM TYPES',   value: stats.items.toLocaleString() },
              ].map((s, i) => (
                <div key={i} style={{
                  background: 'var(--surface)', padding: '20px 24px',
                  borderRadius: i === 0 ? '12px 0 0 12px' : i === 3 ? '0 12px 12px 0' : 0,
                  border: '1px solid var(--border)',
                }}>
                  <div className="mono" style={{ color: 'var(--muted)', fontSize: 10, letterSpacing: 2, marginBottom: 6 }}>
                    {s.label}
                  </div>
                  <div className="mono" style={{ color: 'var(--accent)', fontSize: 22, fontWeight: 700 }}>
                    {s.value}
                  </div>
                </div>
              ))}
            </div>

            {/* ── Output Mode Toggle ── */}
            <div style={{
              display: 'flex', alignItems: 'center', gap: 12,
              marginBottom: 16,
            }}>
              <span className="mono" style={{ fontSize: 11, color: 'var(--muted)', letterSpacing: 2 }}>
                OUTPUT MODE
              </span>
              <div style={{
                display: 'flex',
                background: 'var(--surface)',
                border: '1px solid var(--border)',
                borderRadius: 8,
                overflow: 'hidden',
              }}>
                {(['qty', 'amount'] as OutputMode[]).map((mode) => (
                  <button
                    key={mode}
                    onClick={() => handleModeChange(mode)}
                    style={{
                      padding: '8px 20px',
                      border: 'none',
                      cursor: 'pointer',
                      fontFamily: 'Space Mono, monospace',
                      fontSize: 12,
                      letterSpacing: 1,
                      fontWeight: 600,
                      transition: 'all 0.15s',
                      background: outputMode === mode ? 'var(--accent)' : 'transparent',
                      color:      outputMode === mode ? '#fff'          : 'var(--muted)',
                      borderRight: mode === 'qty' ? '1px solid var(--border)' : 'none',
                    }}
                  >
                    {mode === 'qty' ? '# QUANTITY' : 'PKR AMOUNT'}
                  </button>
                ))}
              </div>
              <span style={{ fontSize: 12, color: 'var(--muted)' }}>
                {outputMode === 'qty'
                  ? 'Item cells show units sold'
                  : 'Item cells show total sale value (PKR)'}
              </span>
            </div>

            {/* Preview table */}
            <div style={{
              background: 'var(--surface)',
              border: '1px solid var(--border)',
              borderRadius: 12,
              overflow: 'hidden',
              marginBottom: 20,
            }}>
              <div style={{
                padding: '12px 20px',
                borderBottom: '1px solid var(--border)',
                display: 'flex', justifyContent: 'space-between', alignItems: 'center',
              }}>
                <span className="mono" style={{ fontSize: 11, color: 'var(--muted)', letterSpacing: 2 }}>
                  PREVIEW (first {preview.length - 1} rows) · {outputMode === 'qty' ? 'QUANTITY' : 'AMOUNT'}
                </span>
                <span style={{ fontSize: 12, color: 'var(--muted)' }}>{fileName}</span>
              </div>
              <div style={{ overflowX: 'auto' }}>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 12 }}>
                  <thead>
                    <tr style={{ background: 'var(--surface2)' }}>
                      {preview[0]?.map((h, i) => (
                        <th key={i} style={{
                          padding: '10px 14px', textAlign: 'left',
                          color: 'var(--accent)', fontWeight: 600,
                          borderBottom: '1px solid var(--border)',
                          whiteSpace: 'nowrap', fontFamily: 'Space Mono, monospace',
                          fontSize: 10, letterSpacing: 1,
                        }}>
                          {String(h)}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {preview.slice(1).map((row, ri) => (
                      <tr key={ri} style={{
                        borderBottom: '1px solid var(--border)',
                        background: ri % 2 === 0 ? 'transparent' : 'rgba(255,255,255,0.01)',
                      }}>
                        {row.map((cell, ci) => (
                          <td key={ci} style={{
                            padding: '9px 14px', color: 'var(--text)',
                            whiteSpace: 'nowrap',
                          }}>
                            {cell === '' || cell === null || cell === undefined ? (
                              <span style={{ color: 'var(--border)' }}>—</span>
                            ) : String(cell)}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Actions */}
            <div style={{ display: 'flex', gap: 12 }}>
              <button
                onClick={downloadXLSX}
                style={{
                  flex: 1, padding: '16px 32px',
                  background: 'var(--accent)', border: 'none',
                  borderRadius: 10, color: '#fff',
                  fontSize: 15, fontWeight: 600, cursor: 'pointer',
                  fontFamily: 'Space Mono, monospace',
                  letterSpacing: 1,
                  transition: 'all 0.2s',
                  boxShadow: '0 4px 20px rgba(249,115,22,0.3)',
                }}
                onMouseOver={e => (e.currentTarget.style.background = 'var(--accent2)')}
                onMouseOut={e => (e.currentTarget.style.background = 'var(--accent)')}
              >
                ↓ DOWNLOAD {outputMode === 'qty' ? 'QUANTITY' : 'AMOUNT'} XLSX
              </button>
              <button
                onClick={reset}
                style={{
                  padding: '16px 24px',
                  background: 'transparent',
                  border: '1px solid var(--border)',
                  borderRadius: 10, color: 'var(--muted)',
                  fontSize: 14, cursor: 'pointer',
                  transition: 'all 0.2s',
                }}
                onMouseOver={e => { e.currentTarget.style.borderColor = 'var(--text)'; e.currentTarget.style.color = 'var(--text)'; }}
                onMouseOut={e => { e.currentTarget.style.borderColor = 'var(--border)'; e.currentTarget.style.color = 'var(--muted)'; }}
              >
                Convert Another
              </button>
            </div>
          </div>
        )}

        {/* Footer */}
        <div style={{
          marginTop: 48, paddingTop: 24,
          borderTop: '1px solid var(--border)',
          display: 'flex', justifyContent: 'space-between', alignItems: 'center',
        }}>
          <span className="mono" style={{ fontSize: 11, color: 'var(--muted)', letterSpacing: 1 }}>
            JT RESTAURANT · SALES TOOLS
          </span>
          <span style={{ fontSize: 12, color: 'var(--border)' }}>
            All processing happens locally in your browser
          </span>
        </div>
      </div>
    </div>
  );
}

function formatDisplayDate(d: string): string {
  if (!d) return '—';
  const m = d.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (m) return `${m[3]}/${m[2]}/${m[1]}`;
  return d;
}
