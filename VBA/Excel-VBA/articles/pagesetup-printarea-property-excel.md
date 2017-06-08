---
title: PageSetup.PrintArea Property (Excel)
keywords: vbaxl10.chm473092
f1_keywords:
- vbaxl10.chm473092
ms.prod: excel
api_name:
- Excel.PageSetup.PrintArea
ms.assetid: da4d5231-cc74-5940-ffd4-224b78e5244c
ms.date: 06/08/2017
---


# PageSetup.PrintArea Property (Excel)

Returns or sets the range to be printed, as a string using A1-style references in the language of the macro. Read/write  **String** .


## Syntax

 _expression_ . **PrintArea**

 _expression_ A variable that represents a **PageSetup** object.


## Remarks

Set this property to  **False** or to the empty string ("") to set the print area to the entire sheet.

This property applies only to worksheet pages.


## Example

This example sets the print area to cells A1:C5 on Sheet1.


```vb
Worksheets("Sheet1").PageSetup.PrintArea = "$A$1:$C$5"
```

This example sets the print area to the current region on Sheet1. Note that you use the  **Address** property to return an A1-style address.




```vb
Worksheets("Sheet1").Activate 
ActiveSheet.PageSetup.PrintArea = _ 
 ActiveCell.CurrentRegion.Address
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-excel.md)

