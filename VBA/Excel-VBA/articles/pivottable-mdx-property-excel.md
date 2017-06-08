---
title: PivotTable.MDX Property (Excel)
keywords: vbaxl10.chm235143
f1_keywords:
- vbaxl10.chm235143
ms.prod: excel
api_name:
- Excel.PivotTable.MDX
ms.assetid: 50a211c9-4b46-568c-5313-fd093d99a140
ms.date: 06/08/2017
---


# PivotTable.MDX Property (Excel)

Returns a  **String** indicating the Multidimensional Expression (MDX) that would be sent to the provider to populate the current PivotTable view. Read-only.


## Syntax

 _expression_ . **MDX**

 _expression_ A variable that represents a **PivotTable** object.


## Remarks

Querying this value for a non-Online Analytical Processing (OLAP) PivotTable, or when there is no PivotTable view (no data items), will return a run-time error.


## Example

This example returns the MDX string for the PivotTable. It assumes that a PivotTable exists on the active worksheet.


```vb
Sub CheckMDX() 
 
 Dim pvtTable As PivotTable 
 
 Set pvtTable = ActiveSheet.PivotTables(1) 
 
 MsgBox "The MDX string for the PivotTable is: " &; _ 
 pvtTable.MDX 
 
End Sub
```


## See also


#### Concepts


[PivotTable Object](pivottable-object-excel.md)

