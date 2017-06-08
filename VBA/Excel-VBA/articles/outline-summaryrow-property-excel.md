---
title: Outline.SummaryRow Property (Excel)
keywords: vbaxl10.chm455076
f1_keywords:
- vbaxl10.chm455076
ms.prod: excel
api_name:
- Excel.Outline.SummaryRow
ms.assetid: f36fac55-cafd-1ec6-4e85-a7f4fc665c04
ms.date: 06/08/2017
---


# Outline.SummaryRow Property (Excel)

Returns or sets the location of the summary rows in the outline. Read/write  **[XlSummaryRow](xlsummaryrow-enumeration-excel.md)** .


## Syntax

 _expression_ . **SummaryRow**

 _expression_ A variable that represents an **Outline** object.


## Remarks



| **XlSummaryRow** can be one of these **XlSummaryRow** constants.|
| **xlSummaryBelow** The summary row will be positioned below the detail rows in the outline.|
| **xlSummaryAbove** The summary row will be positioned above the detail rows in the outline.|
Set  **XlSummaryRow** to **xlAbove** for Microsoft Word-style outlines, where category headers are above the detailed information. Set **XlSummaryRow** to **xlBelow** for accounting-style outlines, where summations are below the detailed information.


## Example

This example creates an outline with automatic styles, with the summary row above the detail rows, and with the summary column to the right of the detail columns.


```vb
Worksheets("Sheet1").Activate 
Selection.AutoOutline 
With ActiveSheet.Outline 
 .SummaryRow = xlAbove 
 .SummaryColumn = xlRight 
 .AutomaticStyles = True 
End With
```


## See also


#### Concepts


[Outline Object](outline-object-excel.md)

