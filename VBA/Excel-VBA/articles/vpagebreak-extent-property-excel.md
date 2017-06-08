---
title: VPageBreak.Extent Property (Excel)
keywords: vbaxl10.chm156077
f1_keywords:
- vbaxl10.chm156077
ms.prod: excel
api_name:
- Excel.VPageBreak.Extent
ms.assetid: 5ad74c42-16dd-86d4-723f-d8d316b70285
ms.date: 06/08/2017
---


# VPageBreak.Extent Property (Excel)

Returns the type of the specified page break: full-screen or only within a print area. Can be either of the following  **[XlPageBreakExtent](xlpagebreakextent-enumeration-excel.md)** constants: **xlPageBreakFull** or **xlPageBreakPartial** . Read-only **Long** .


## Syntax

 _expression_ . **Extent**

 _expression_ A variable that represents a **VPageBreak** object.


## Example

This example displays the total number of full-screen and print-area horizontal page breaks.


```vb
For Each pb in Worksheets(1).HPageBreaks 
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


[VPageBreak Object](vpagebreak-object-excel.md)

