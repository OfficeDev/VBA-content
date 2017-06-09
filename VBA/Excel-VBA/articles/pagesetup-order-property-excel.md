---
title: PageSetup.Order Property (Excel)
keywords: vbaxl10.chm473089
f1_keywords:
- vbaxl10.chm473089
ms.prod: excel
api_name:
- Excel.PageSetup.Order
ms.assetid: cb39c83a-3ab2-cab9-531c-762db4ab6770
ms.date: 06/08/2017
---


# PageSetup.Order Property (Excel)

Returns or sets a  **[XlOrder](xlorder-enumeration-excel.md)** value that represents the order that Microsoft Excel uses to number pages when printing a large worksheet.


## Syntax

 _expression_ . **Order**

 _expression_ A variable that represents a **PageSetup** object.


## Example

This example breaks Sheet1 into pages when the worksheet is printed. Numbering and printing proceed from the first page to the pages to the right, and then move down and continue printing across the sheet.


```vb
Worksheets("Sheet1").PageSetup.Order = xlOverThenDown
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-excel.md)

