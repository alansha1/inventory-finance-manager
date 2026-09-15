# Inventory & Finance Manager (Advanced Excel)

A combined inventory tracking + financial modeling workbook: real-time stock status
with data validation, a budget-vs-actual finance model, Power Query/Power Pivot data
relationships, and a VBA macro that automates purchase-order creation for low-stock items.

## What's inside

- **`Lists`** — hidden helper sheet holding the dropdown lists (Categories, Suppliers).
- **`Suppliers`** — supplier master table (`SupplierData`): ID, name, contact, lead time,
  payment terms.
- **`Inventory`** — 54-item stock table (`InventoryData`) with **data validation**
  dropdowns for Category and Supplier, and formula-driven columns: Inventory Value,
  Reorder Cost, and a live `Status` flag (`IF(QuantityOnHand<=ReorderLevel,"Reorder","OK")`).
- **`PurchaseOrders`** — 60-row PO log (`POData`) that looks up item details from
  `Inventory` with `INDEX`/`MATCH`, and calculates cost per order.
- **`BudgetModel`** — editable monthly budget by category (blue/yellow input cells),
  actual spend pulled live via `SUMIFS` from Purchase Orders, variance, and a total
  forecast reorder cost.
- **`Dashboard`** — KPI cards, a category-spend bar chart, and a monthly spend trend chart.

## Skills demonstrated

Data validation (dropdown lists) · `INDEX`/`MATCH` lookups · `SUMIFS`/`SUMIF`/`COUNTIF` ·
Excel Tables · financial modeling (budget vs. actual, variance) · KPI dashboard design ·
VBA macros · Power Query · Power Pivot data modeling

## Parts you add yourself in Excel

Some pieces are genuinely Excel-app-only and are worth building hands-on:

### 1. VBA macro — auto-generate POs for low-stock items

1. Save the file as `.xlsm` (File → Save As → Excel Macro-Enabled Workbook).
2. Press **Alt+F11** to open the VBA editor → **Insert → Module**.
3. Paste in this macro:

   ```vba
   Sub GenerateReorderPOs()
       Dim invSheet As Worksheet, poSheet As Worksheet
       Dim invTable As ListObject, poTable As ListObject
       Dim i As Long, newRow As ListRow
       Dim itemID As Variant, qty As Variant, count As Long

       Set invSheet = ThisWorkbook.Sheets("Inventory")
       Set poSheet = ThisWorkbook.Sheets("PurchaseOrders")
       Set invTable = invSheet.ListObjects("InventoryData")
       Set poTable = poSheet.ListObjects("POData")
       count = 0

       For i = 1 To invTable.ListRows.Count
           If invTable.DataBodyRange(i, 12).Value = "Reorder" Then
               Set newRow = poTable.ListRows.Add
               newRow.Range(1, 1).Value = 7000 + poTable.ListRows.Count
               newRow.Range(1, 2).Value = Date
               newRow.Range(1, 3).Value = invTable.DataBodyRange(i, 1).Value
               newRow.Range(1, 7).Value = invTable.DataBodyRange(i, 9).Value
               count = count + 1
           End If
       Next i

       MsgBox count & " purchase order(s) generated for low-stock items.", vbInformation
   End Sub
   ```

4. Close the VBA editor. On the `Dashboard` sheet, go to **Developer → Insert → Button
   (Form Control)**, draw it near the note in A10, and assign it to `GenerateReorderPOs`.
   Label it "Generate POs for Low Stock".
5. Click the button — it scans `Inventory` for every item flagged `Reorder` and appends
   a new row to `PurchaseOrders` automatically.

### 2. Power Query — consolidate supplier data

1. **Data → Get Data → From Table/Range**, select `InventoryData`, and load it into
   Power Query.
2. Merge it with `SupplierData` on the `Supplier`/`SupplierName` columns (Merge Queries)
   to pull `LeadTimeDays` and `PaymentTerms` alongside each inventory item.
3. Load the merged result as a new table, e.g. `InventoryWithSupplierInfo`.

### 3. Power Pivot — build the data model

1. **Insert → PivotTable → From Data Model** (or Power Pivot tab → Add to Data Model)
   for `InventoryData`, `POData`, and `SupplierData`.
2. Create relationships: `InventoryData[Supplier]` ↔ `SupplierData[SupplierName]`,
   and `POData[ItemID]` ↔ `InventoryData[ItemID]`.
3. Build a PivotTable from the model showing Total PO Spend by Supplier and Lead Time —
   this is only possible once the relationships exist, and is the payoff of the data model.

## Dataset

Synthetic: 54 inventory items across 6 categories, 6 suppliers, and 60 historical
purchase orders over the last 6 months. Generated for demonstration purposes.

## Tools used

Microsoft Excel (VBA, Power Query, Power Pivot) · openpyxl (base workbook generation)
## Proof of automation

   ![VBA automation](vba-automation.png)
   Clicking "Generate POs for Low Stock" scans Inventory and auto-creates purchase orders for every item below its reorder level.

   ![Power Pivot data model](power-pivot-model.png)
   Relationships connecting Inventory, Purchase Orders, and Suppliers — enabling cross-table analysis without manual lookups.
