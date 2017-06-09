---
title: PageSetup.PrintComments Property (Excel)
keywords: vbaxl10.chm473104
f1_keywords:
- vbaxl10.chm473104
ms.prod: excel
api_name:
- Excel.PageSetup.PrintComments
ms.assetid: 1f479032-ca02-982f-5877-83c776ce2611
ms.date: 06/08/2017
---


# PageSetup.PrintComments Property (Excel)

Returns or sets the way comments are printed with the sheet. Read/write  **[XlPrintLocation](xlprintlocation-enumeration-excel.md)** .


## Syntax

 _expression_ . **PrintComments**

 _expression_ A variable that represents a **PageSetup** object.


## Remarks





| **XlPrintLocation** can be one of these **XlPrintLocation** constants.|
| **xlPrintInPlace**|
| **xlPrintNoComments**|
| **xlPrintSheetEnd**|

## Example

This example causes comments to be printed as end notes when worksheet one is printed.


```vb
Worksheets(1).PageSetup.PrintComments = xlPrintSheetEnd
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-excel.md)

