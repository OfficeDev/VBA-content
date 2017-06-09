---
title: PageSetup.PrintHeadings Property (Excel)
keywords: vbaxl10.chm473094
f1_keywords:
- vbaxl10.chm473094
ms.prod: excel
api_name:
- Excel.PageSetup.PrintHeadings
ms.assetid: 027441c6-da40-f518-a166-adb54da02a27
ms.date: 06/08/2017
---


# PageSetup.PrintHeadings Property (Excel)

 **True** if row and column headings are printed with this page. Applies only to worksheets. Read/write **Boolean** .


## Syntax

 _expression_ . **PrintHeadings**

 _expression_ A variable that represents a **PageSetup** object.


## Remarks

The  **[DisplayHeadings](window-displayheadings-property-excel.md)** property controls the on-screen display of headings.


## Example

This example turns off the printing of headings for Sheet1.


```vb
Worksheets("Sheet1").PageSetup.PrintHeadings = False
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-excel.md)

