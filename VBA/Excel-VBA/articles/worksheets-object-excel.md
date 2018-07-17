---
title: Worksheets Object (Excel)
keywords: vbaxl10.chm469072
f1_keywords:
- vbaxl10.chm469072
ms.prod: excel
api_name:
- Excel.Worksheets
ms.assetid: 5ec467a6-97e3-98d7-0b14-845d20c15910
ms.date: 06/08/2017
---


# Worksheets Object (Excel)

A collection of all the  **[Worksheet](worksheet-object-excel.md)** objects in the specified or active workbook. Each **Worksheet** object represents a worksheet.

## Remarks

The  **Worksheet** object is also a member of the [Sheets](sheets-object-excel.md) collection. The **Sheets** collection contains all the sheets in the workbook (both chart sheets and worksheets).

## Example

Use the  **[Worksheets](workbook-worksheets-property-excel.md)** property to return the **Worksheets** collection.The following example moves all the worksheets to the end of the workbook.

```
Worksheets.Move After:=Sheets(Sheets.Count)
```

Use the  **[Add](worksheets-add-method-excel.md)** method to create a new worksheet and add it to the collection. The following example adds two new worksheets before sheet one of the active workbook.

```
Worksheets.Add Count:=2, Before:=Sheets(1)
```

Use  **Worksheets** ( _index_ ), where _index_ is the worksheet index number or name, to return a single **Worksheet** object. The following example hides worksheet one in the active workbook.

```
Worksheets(1).Visible = False
```
## Methods

|**Name**|
|:-----|
|[Add](worksheets-add-method-excel.md)|
|[Add2](worksheets-add2-method-excel.md)|
|[Copy](worksheets-copy-method-excel.md)|
|[Delete](worksheets-delete-method-excel.md)|
|[FillAcrossSheets](worksheets-fillacrosssheets-method-excel.md)|
|[Move](worksheets-move-method-excel.md)|
|[PrintOut](worksheets-printout-method-excel.md)|
|[PrintPreview](worksheets-printpreview-method-excel.md)|
|[Select](worksheets-select-method-excel.md)|

## Properties

|**Name**|
|:-----|
|[Application](worksheets-application-property-excel.md)|
|[Count](worksheets-count-property-excel.md)|
|[Creator](worksheets-creator-property-excel.md)|
|[HPageBreaks](worksheets-hpagebreaks-property-excel.md)|
|[Item](worksheets-item-property-excel.md)|
|[Parent](worksheets-parent-property-excel.md)|
|[Visible](worksheets-visible-property-excel.md)|
|[VPageBreaks](worksheets-vpagebreaks-property-excel.md)|

## See also

#### Other resources

[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
