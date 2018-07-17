---
title: PageSetup.FitToPagesWide Property (Excel)
keywords: vbaxl10.chm473083
f1_keywords:
- vbaxl10.chm473083
ms.prod: excel
api_name:
- Excel.PageSetup.FitToPagesWide
ms.assetid: 162bd2d2-35fa-8133-ab1c-27dcfc173317
ms.date: 06/08/2017
---


# PageSetup.FitToPagesWide Property (Excel)

Returns or sets the number of pages wide the worksheet will be scaled to when it's printed. Applies only to worksheets. Read/write  **Variant** .


## Syntax

 _expression_ . **FitToPagesWide**

 _expression_ A variable that represents a **PageSetup** object.


## Remarks

If this property is  **False** , Microsoft Excel scales the worksheet according to the **[FitToPagesTall](pagesetup-fittopagestall-property-excel.md)** property.

If the  **[Zoom](pagesetup-zoom-property-excel.md)** property is **True** , the **FitToPagesWide** property is ignored.


## Example

This example causes Microsoft Excel to print Sheet1 exactly one page wide and tall.


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

