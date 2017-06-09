---
title: TableStyleElement.Borders Property (Excel)
keywords: vbaxl10.chm835075
f1_keywords:
- vbaxl10.chm835075
ms.prod: excel
api_name:
- Excel.TableStyleElement.Borders
ms.assetid: a6fdfe85-0953-f796-5c89-6f418e9226e6
ms.date: 06/08/2017
---


# TableStyleElement.Borders Property (Excel)

Returns a  **[Borders](borders-object-excel.md)** collection that represents the borders of a table style element. Read-only.


## Syntax

 _expression_ . **Borders**

 _expression_ A variable that represents a **TableStyleElement** object.


## Example

This example sets the color of the top border of a table to red.


```vb
With ActiveWorkbook.TableStyles("Table Style 4").TableStyleElements( _ 
 xlWholeTable).Borders(xlEdgeTop) 
 .Color = 255 
 .TintAndShade = 0 
 .Weight = 2 
 .LineStyle = 1 
End With
```


## See also


#### Concepts


[TableStyleElement Object](tablestyleelement-object-excel.md)

