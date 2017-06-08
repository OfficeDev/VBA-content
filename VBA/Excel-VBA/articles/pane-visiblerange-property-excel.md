---
title: Pane.VisibleRange Property (Excel)
keywords: vbaxl10.chm360079
f1_keywords:
- vbaxl10.chm360079
ms.prod: excel
api_name:
- Excel.Pane.VisibleRange
ms.assetid: 03853894-ca83-1672-21bb-15099bab03d8
ms.date: 06/08/2017
---


# Pane.VisibleRange Property (Excel)

Returns a  **[Range](range-object-excel.md)** object that represents the range of cells that are visible in the window or pane. If a column or row is partially visible, it's included in the range. Read-only.


## Syntax

 _expression_ . **VisibleRange**

 _expression_ A variable that represents a **Pane** object.


## Example

This example displays the number of cells visible on Sheet1.


```vb
Worksheets("Sheet1").Activate 
MsgBox "There are " &; Windows(1).VisibleRange.Cells.Count _ 
 &; " cells visible"
```


## See also


#### Concepts


[Pane Object](pane-object-excel.md)

