---
title: Worksheets.VPageBreaks Property (Excel)
keywords: vbaxl10.chm470085
f1_keywords:
- vbaxl10.chm470085
ms.prod: excel
api_name:
- Excel.Worksheets.VPageBreaks
ms.assetid: 09c097f5-6344-ea88-2ce4-a582f84f2fe5
ms.date: 06/08/2017
---


# Worksheets.VPageBreaks Property (Excel)

Returns a  **[VPageBreaks](worksheets-vpagebreaks-property-excel.md)** collection that represents the vertical page breaks on the sheet. Read-only.


## Syntax

 _expression_ . **VPageBreaks**

 _expression_ A variable that represents a **Worksheets** object.


## Example

This example displays the total number of full-screen and print-area vertical page breaks.


```vb
For Each pb in Worksheets(1).VPageBreaks 
 If pb.Extent = xlPageBreakFull Then 
 cFull = cFull + 1 
 Else 
 cPartial = cPartial + 1 
 End If 
Next 
MsgBox cFull &; " full-screen page breaks, " &; cPartial &; _ 
 " print-area page breaks"
```


## See also


#### Concepts


[Worksheets Object](worksheets-object-excel.md)

