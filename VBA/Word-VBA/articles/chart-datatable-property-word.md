---
title: Chart.DataTable Property (Word)
ms.prod: word
api_name:
- Word.Chart.DataTable
ms.assetid: 1cae3588-5bc4-5ec4-c3f3-cc642d0755a6
ms.date: 06/08/2017
---


# Chart.DataTable Property (Word)

Returns the chart data table. Read-only  **[DataTable](datatable-object-word.md)** .


## Syntax

 _expression_ . **DataTable**

 _expression_ A variable that represents a **[Chart](chart-object-word.md)** object.


## Example

The following example adds a data table with an outline border to the first chart in the active document.


```vb
With ActiveDocument.InlineShapes(1) 
 If .HasChart Then 
 .Chart.HasDataTable = True 
 .Chart.DataTable.HasBorderOutline = True 
 End If 
End With
```


## See also


#### Concepts


[Chart Object](chart-object-word.md)

