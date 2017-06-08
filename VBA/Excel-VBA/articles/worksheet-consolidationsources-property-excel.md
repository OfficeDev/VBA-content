---
title: Worksheet.ConsolidationSources Property (Excel)
keywords: vbaxl10.chm175089
f1_keywords:
- vbaxl10.chm175089
ms.prod: excel
api_name:
- Excel.Worksheet.ConsolidationSources
ms.assetid: d7868b1c-c9ae-97c5-a092-033fe52db5d4
ms.date: 06/08/2017
---


# Worksheet.ConsolidationSources Property (Excel)

Returns an array of string values that name the source sheets for the worksheet's current consolidation. Returns  **Empty** if there's no consolidation on the sheet. Read-only **Variant** .


## Syntax

 _expression_ . **ConsolidationSources**

 _expression_ A variable that represents a **Worksheet** object.


## Example

This example displays the names of the source ranges for the consolidation on Sheet1. The list appears on a new worksheet created by the example.


```vb
Set newSheet = Worksheets.Add 
newSheet.Range("A1").Value = "Consolidation Sources" 
aSources = Worksheets("Sheet1").ConsolidationSources 
If IsEmpty(aSources) Then 
 newSheet.Range("A2").Value = "none" 
Else 
 For i = 1 To UBound(aSources) 
 newSheet.Cells(i + 1, 1).Value = aSources(i) 
 Next i 
End If 
newSheet.Columns("A:B").AutoFit
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

