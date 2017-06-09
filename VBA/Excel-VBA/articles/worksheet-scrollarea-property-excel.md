---
title: Worksheet.ScrollArea Property (Excel)
keywords: vbaxl10.chm175124
f1_keywords:
- vbaxl10.chm175124
ms.prod: excel
api_name:
- Excel.Worksheet.ScrollArea
ms.assetid: 7421676d-3a98-3826-31f9-80e7c8946777
ms.date: 06/08/2017
---


# Worksheet.ScrollArea Property (Excel)

Returns or sets the range where scrolling is allowed, as an A1-style range reference. Cells outside the scroll area cannot be selected. Read/write  **String** .


## Syntax

 _expression_ . **ScrollArea**

 _expression_ A variable that represents a **Worksheet** object.


## Remarks

Set this property to the empty string ("") to enable cell selection for the entire sheet.


## Example

This example sets the scroll area for worksheet one.


```vb
Worksheets(1).ScrollArea = "a1:f10"
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

