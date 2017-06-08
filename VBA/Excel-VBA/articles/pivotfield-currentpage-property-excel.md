---
title: PivotField.CurrentPage Property (Excel)
keywords: vbaxl10.chm240077
f1_keywords:
- vbaxl10.chm240077
ms.prod: excel
api_name:
- Excel.PivotField.CurrentPage
ms.assetid: 4a59fe58-8f95-4cf3-d4a3-ab6ea6b24b8a
ms.date: 06/08/2017
---


# PivotField.CurrentPage Property (Excel)

Returns or sets the current page showing for the page field (valid only for page fields). Read/write  **[PivotItem](pivotitem-object-excel.md)** .


## Syntax

 _expression_ . **CurrentPage**

 _expression_ A variable that represents a **PivotField** object.


## Example

This example returns the current page name for the PivotTable report on Sheet1 in the string variable  `strPgName`.


```vb
Set pvtTable = Worksheets("Sheet1").Range("A3").PivotTable 
strPgName = pvtTable.PivotFields("Country").CurrentPage.Name
```


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

