---
title: PivotField.CurrentPageList Property (Excel)
keywords: vbaxl10.chm240135
f1_keywords:
- vbaxl10.chm240135
ms.prod: excel
api_name:
- Excel.PivotField.CurrentPageList
ms.assetid: 3efde5a2-4cf3-b95d-e7ba-65ea8e184e64
ms.date: 06/08/2017
---


# PivotField.CurrentPageList Property (Excel)

Returns or sets an array of strings corresponding to the list of items included in a multiple-item page field of a PivotTable report. Read/write  **Variant** .


## Syntax

 _expression_ . **CurrentPageList**

 _expression_ A variable that represents a **PivotField** object.


## Remarks

To avoid run-time errors, the data source must be an OLAP source, the field chosen must currently be in the Page position, and the  **[EnableMultiplePageItems](pivotfield-enablemultiplepageitems-property-excel.md)** property must be set to **True** .


## Example

This example sets the page field to list the "Food" items of the PivotTable report. It assumes a PivotTable exists on the active worksheet.


```vb
Sub UseCurrentPageList() 
 
 Dim pvtTable As PivotTable 
 Dim pvtField As PivotField 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 Set pvtField = pvtTable.PivotFields("[Product]") 
 
 ' To avoid run-time errors set the following property to True. 
 pvtTable.CubeFields("[Product]").EnableMultiplePageItems = True 
 
 ' Set the page list to "Food". 
 pvtField.CurrentPageList = "[Product].[All Products].[Food]" 
 
End Sub
```


## See also


#### Concepts


[PivotField Object](pivotfield-object-excel.md)

