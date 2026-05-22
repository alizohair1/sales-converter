# JT Sales Register Converter

A cloud-based web app that converts your **Sales Register** (`.xls`/`.xlsx`) into a **10-Minute Interval Summary** — the same format as `sales_10min_06apr_jt.xlsx`.

## What it does

1. You upload your daily `sr_XX_XX_jt.xls` file
2. The app parses all orders, items, payment methods, customer types, and tills
3. It groups everything into 10-minute time buckets
4. You download a clean `.xlsx` with one row per [date × time × till × customer × payment] combination, with item quantities as columns

## Stack

- **Next.js 14** (App Router)
- **SheetJS (xlsx)** for reading `.xls` / `.xlsx` and writing output
- **FileSaver.js** for download
- **Tailwind CSS** for styling
- Deployed on **Vercel** (zero config)

## Deploy in 3 steps

### 1. Push to GitHub

```bash
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/YOUR_USERNAME/sales-converter.git
git push -u origin main
```

### 2. Connect to Vercel

1. Go to [vercel.com](https://vercel.com) and sign in with GitHub
2. Click **"Add New Project"**
3. Import your `sales-converter` repository
4. Leave all settings as default (Vercel auto-detects Next.js)
5. Click **Deploy**

That's it. Vercel gives you a live URL like `https://sales-converter-abc.vercel.app`.

### 3. Use it

Open the URL in any browser, drag-drop your `.xls` sales register, and download the 10-min interval Excel file.

## Local development

```bash
npm install
npm run dev
# Open http://localhost:3000
```

## How the parsing works

The Sales Register has this structure per order:

```
DATE : 2026-04-06   A  12.0  :  1  TILL :  TAKE AWAY  CUSTOMER :  CASH CUSTOMER  PAYMENT :  CASH
  [item rows: code | name | qty | rate | amount]
  [summary: G.AMOUNT : gross  DIS :  N.AMOUNT : net]
```

The converter:
1. Detects each `DATE :` row as an order header
2. Collects all item rows beneath it
3. Reads the `G.AMOUNT` summary row
4. Groups orders into 10-min buckets by [date + time + till + customer + payment]
5. Pivots items into columns

## Privacy

All file processing happens **100% in the browser** using JavaScript. Your sales data never leaves your device and is never sent to any server.
