---
title: PageSetup.FooterMargin Property (Excel)
keywords: vbaxl10.chm473084
f1_keywords:
- vbaxl10.chm473084
ms.prod: excel
api_name:
- Excel.PageSetup.FooterMargin
ms.assetid: b6ec4b9c-c828-e6fe-2a65-ccddd1b05c30
ms.date: 06/08/2017
---


# PageSetup.FooterMargin Property (Excel)

Returns or sets the distance from the bottom of the page to the footer, in points. Read/write  **Double** .


## Syntax

 _expression_ . **FooterMargin**

 _expression_ A variable that represents a **PageSetup** object.


## Example

This example sets the footer margin of Sheet1 to 0.5 inch.


```vb
Worksheets("Sheet1").PageSetup.FooterMargin = _ 
 Application.InchesToPoints(0.5)
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-excel.md)

