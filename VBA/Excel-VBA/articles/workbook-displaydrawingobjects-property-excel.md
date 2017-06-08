---
title: Workbook.DisplayDrawingObjects Property (Excel)
keywords: vbaxl10.chm199098
f1_keywords:
- vbaxl10.chm199098
ms.prod: excel
api_name:
- Excel.Workbook.DisplayDrawingObjects
ms.assetid: 78eec8af-416d-88e6-d1f4-0b97a008f752
ms.date: 06/08/2017
---


# Workbook.DisplayDrawingObjects Property (Excel)

Returns or sets how shapes are displayed. Read/write  **Long** .


## Syntax

 _expression_ . **DisplayDrawingObjects**

 _expression_ A variable that represents a **Workbook** object.


## Remarks





|**Constant**|**Description**|
|:-----|:-----|
| **xlDisplayShapes**|Show all shapes.|
| **xlPlaceholders**|Show only placeholders.|
| **xlHide**|Hide all shapes.|

## Example

This example hides all the shapes in the active workbook.


```vb
ActiveWorkbook.DisplayDrawingObjects = xlHide
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

