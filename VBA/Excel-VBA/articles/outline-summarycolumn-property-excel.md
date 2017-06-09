---
title: Outline.SummaryColumn Property (Excel)
keywords: vbaxl10.chm455075
f1_keywords:
- vbaxl10.chm455075
ms.prod: excel
api_name:
- Excel.Outline.SummaryColumn
ms.assetid: b134c991-7875-445a-ca73-d48bf23f3eea
ms.date: 06/08/2017
---


# Outline.SummaryColumn Property (Excel)

Returns or sets the location of the summary columns in the outline. Read/write  **[XlSummaryColumn](xlsummarycolumn-enumeration-excel.md)** .


## Syntax

 _expression_ . **SummaryColumn**

 _expression_ A variable that represents an **Outline** object.


## Remarks





| **XlSummaryColumn** can be one of these **XlSummaryColumn** constants.|
| **xlSummaryOnRight** The summary column will be positioned to the right of the detail columns in the outline.|
| **xlSummaryOnLeft** The summary column will be positioned to the left of the detail columns in the outline.|

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

