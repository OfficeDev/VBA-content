---
title: PivotTable.PivotSelectionStandard Property (Excel)
keywords: vbaxl10.chm235138
f1_keywords:
- vbaxl10.chm235138
ms.prod: excel
api_name:
- Excel.PivotTable.PivotSelectionStandard
ms.assetid: 72252681-65ec-885b-466d-fb890db812a4
ms.date: 06/08/2017
---


# PivotTable.PivotSelectionStandard Property (Excel)

Returns or sets a  **String** indicating the PivotTable selection in standard PivotTable report format using English (United States) settings. Read/write.


## Syntax

 _expression_ . **PivotSelectionStandard**

 _expression_ A variable that represents a **PivotTable** object.


## Remarks

The  **PivotSelectionStandard** property is "international-friendly" whereas the **PivotSelection** method is not.


## Example

This example selects a field titled "1.57" in the PivotTable and inserts a blank column field before it. The example assumes a PivotTable exists on the active worksheet that contains a column field titled "1.57".


```vb
Sub CheckPivotSelectionStandard() 
 
 Dim pvtTable As PivotTable 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 pvtTable.PivotSelectionStandard = "1.57" 
 Selection.Insert 
 
End Sub
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

