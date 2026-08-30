# Inventory Management System

A single-workbook inventory system built in Excel. It tracks products, vendors, and
customers, logs purchases and sales, and automatically rolls everything up into a live
stock ledger, a P&L snapshot, and low-stock re-order alerts.

## What's in the workbook

| Sheet | Purpose |
|---|---|
| **Dashboard** | Landing page. Shows a notifications panel that lists any product needing re-order, pulled live from the Inventory sheet. |
| **Customers** | Master list of customers: `Cust_ID`, Name, Email, Address. |
| **Vendors** | Master list of vendors per product: `HSN Code`, Product Name, Vendor Name, Phone, Address. |
| **Products** | Product catalog: `HSN Code`, Product Name, Cost (buy price), Selling Price. |
| **New Entry** | Quick-navigation page with links to jump to the Purchase Entry and Sales Entry sheets. |
| **Purchase** | Purchase log — every stock-in transaction (date, vendor, units, cost, amount). |
| **Sales** | Sales log — every stock-out transaction (date, customer, units, price, amount). |
| **Inventory** | The core stock ledger. Calculates current stock per product from Purchase − Sales, values it, and flags re-orders. |
| **pivot** | Pivot tables summarizing sales by customer, units sold by product, stock by product, purchase/sales totals, and overall profit/loss. |

Each data sheet (Customers, Vendors, Products, Purchase, Sales, Inventory) is built as an
**Excel Table**, so ranges auto-expand as rows are added and formulas reference table/column
names (e.g. `purchase[Units]`) instead of cell ranges.

## How the data flows

```
Products  ─┐
Vendors   ─┼──▶  Inventory  ◀──  Purchase (stock in)
Customers ─┘         ▲            Sales     (stock out)
                      │
                  Dashboard  ◀── low-stock notifications
                      │
                    pivot   ◀── summaries & P/L
```

1. **Products / Vendors / Customers** are master data — set these up first, since every
   other sheet looks values up from them via `HSN Code` or `Cust_ID`.
2. **Purchase** and **Sales** are transaction logs. Each row you add automatically pulls
   in the product name, vendor/customer name, and unit cost/price via `VLOOKUP`, then
   calculates the line `Amount`.
3. **Inventory** recalculates per product:
   - `P Units` = total units purchased (`SUMIF` on Purchase)
   - `S Units` = total units sold (`SUMIF` on Sales)
   - `Stock` = `P Units − S Units`
   - `Stock Amt.` = `Stock × Cost`
   - `notifications` = an alert message (with the vendor's phone number) whenever
     `Stock < 5`
4. **Dashboard** surfaces those notification messages so low-stock items are visible at
   a glance.
5. **pivot** summarizes the logs: sales by customer, units sold by product, stock by
   product, total purchase amount, total sales amount, total stock value, and a
   profit/loss figure (Sales Amt. − Purchase Amt.).

## How to use it

**Adding a new product**
1. Add a row to **Products** (HSN Code, Name, Cost, Selling Price).
2. Add a matching row to **Vendors** (same HSN Code, vendor details).
3. Add a row to **Inventory** with the same HSN Code — the formulas will pick up the
   rest automatically.

**Recording a purchase (stock in)**
- Go to **Purchase** (or use the "➡️ Purchase Entry" link on **New Entry**) and add a row:
  HSN Code, Date, Units, Cost. Product name, vendor, and Amount fill in automatically.

**Recording a sale (stock out)**
- Go to **Sales** (or use the "➡️ Sales Entry" link on **New Entry**) and add a row:
  Cust_ID, HSN Code, Date, Units. Customer name, product name, current stock, price,
  and Amount fill in automatically.

**Checking stock / re-orders**
- Open **Inventory** for the full per-product breakdown, or **Dashboard** for a quick
  list of items currently below the re-order threshold (stock < 5 units).

**Refreshing the pivot tables**
- After adding new Purchase/Sales rows, refresh the pivot tables (right-click a pivot
  table → *Refresh*, or *Data → Refresh All*) so the **pivot** sheet reflects the latest
  numbers.

## Key formulas used

- `VLOOKUP` — pulls product/vendor/customer details into the Purchase and Sales logs.
- `SUMIF` — totals purchased/sold units per product for the Inventory sheet.
- `IFERROR` — keeps cells blank instead of showing errors when a lookup has no match yet.
- `IF` + text concatenation — builds the "Needs To Re-Order" alert message with the
  vendor's phone number when stock drops below 5 units.

## Notes & assumptions

- Re-order threshold is hardcoded at **5 units** in the Inventory sheet's notification
  formula — change the `<5` condition there if you want a different threshold.
- Vendor phone numbers in the sample data are partially masked (`98xxxxxxxx`) — replace
  with real numbers for production use.
- All amounts are in ₹ (Indian Rupees), consistent with the HSN Code fields used
  (Indian tax classification codes).
- The workbook has no VBA/macros; all automation is formula-driven, so it works in
  Excel Online, Google Sheets (with VLOOKUP support), and LibreOffice Calc as well.
