---
title: PageSetup.FitToPagesTall Property (Excel)
keywords: vbaxl10.chm473082
f1_keywords:
- vbaxl10.chm473082
ms.prod: excel
api_name:
- Excel.PageSetup.FitToPagesTall
ms.assetid: 1a0141cb-a665-caf5-6bd6-b037f65486dc
ms.date: 06/08/2017
---


# PageSetup.FitToPagesTall Property (Excel)

Returns or sets the number of pages tall the worksheet will be scaled to when it's printed. Applies only to worksheets. Read/write  **Variant** .


## Syntax

 _expression_ . **FitToPagesTall**

 _expression_ A variable that represents a **PageSetup** object.


## Remarks

If this property is  **False** , Microsoft Excel scales the worksheet according to the **[FitToPagesWide](pagesetup-fittopageswide-property-excel.md)** property.

If the  **[Zoom](pagesetup-zoom-property-excel.md)** property is **True** , the **FitToPagesTall** property is ignored.


## Example

This example causes Microsoft Excel to print Sheet1 exactly one page tall and wide.


```vb
With Worksheets("Sheet1").PageSetup 
 .Zoom = False 
 .FitToPagesTall = 1 
 .FitToPagesWide = 1 
End With
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-excel.md)

