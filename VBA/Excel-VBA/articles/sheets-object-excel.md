---
title: Sheets Object (Excel)
keywords: vbaxl10.chm151072
f1_keywords:
- vbaxl10.chm151072
ms.prod: excel
api_name:
- Excel.Sheets
ms.assetid: 048fd93c-bc27-4b58-358f-56fcee1710f8
ms.date: 06/08/2017
---


# Sheets Object (Excel)

A collection of all the sheets in the specified or active workbook.


## Remarks

 The **Sheets** collection can contain **[Chart](chart-object-excel.md)** or **[Worksheet](worksheet-object-excel.md)** objects.

The  **Sheets** collection is useful when you want to return sheets of any type. If you need to work with sheets of only one type, see the object topic for that sheet type.


## Example

Use the  **[Sheets](workbook-sheets-property-excel.md)** property to return the **Sheets** collection. The following example prints all sheets in the active workbook.


```
Sheets.PrintOut
```

Use the  **[Add](sheets-add-method-excel.md)** method to create a new sheet and add it to the collection. The following example adds two chart sheets to the active workbook, placing them after sheet two in the workbook.




```
Sheets.Add type:=xlChart, count:=2, after:=Sheets(2)
```

Use  **Sheets** ( _index_ ), where _index_ is the sheet name or index number, to return a single **Chart** or **Worksheet** object. The following example activates the sheet named "sheet1."




```
Sheets("sheet1").Activate
```

Use  **Sheets** ( _array_ ) to specify more than one sheet. The following example moves the sheets named "Sheet4" and "Sheet5" to the beginning of the workbook.




```
Sheets(Array("Sheet4", "Sheet5")).Move before:=Sheets(1)
```


## Methods



|**Name**|
|:-----|
|[Add](sheets-add-method-excel.md)|
|[Add2](sheets-add2-method-excel.md)|
|[Copy](sheets-copy-method-excel.md)|
|[Delete](sheets-delete-method-excel.md)|
|[FillAcrossSheets](sheets-fillacrosssheets-method-excel.md)|
|[Move](sheets-move-method-excel.md)|
|[PrintOut](sheets-printout-method-excel.md)|
|[PrintPreview](sheets-printpreview-method-excel.md)|
|[Select](sheets-select-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](sheets-application-property-excel.md)|
|[Count](sheets-count-property-excel.md)|
|[Creator](sheets-creator-property-excel.md)|
|[HPageBreaks](sheets-hpagebreaks-property-excel.md)|
|[Item](sheets-item-property-excel.md)|
|[Parent](sheets-parent-property-excel.md)|
|[Visible](sheets-visible-property-excel.md)|
|[VPageBreaks](sheets-vpagebreaks-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
