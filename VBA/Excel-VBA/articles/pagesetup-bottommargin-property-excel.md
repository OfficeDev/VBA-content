---
title: PageSetup.BottomMargin Property (Excel)
keywords: vbaxl10.chm473074
f1_keywords:
- vbaxl10.chm473074
ms.prod: excel
api_name:
- Excel.PageSetup.BottomMargin
ms.assetid: 4c1cd3e0-0ba6-9d2d-4d5a-69d9ee811704
ms.date: 06/08/2017
---


# PageSetup.BottomMargin Property (Excel)

Returns or sets the size of the bottom margin, in points. Read/write  **Double** .


## Syntax

 _expression_ . **BottomMargin**

 _expression_ A variable that represents a **PageSetup** object.


## Remarks

Margins are set or returned in points. Use either the  **[InchesToPoints](application-inchestopoints-method-excel.md)** method or the **[CentimetersToPoints](application-centimeterstopoints-method-excel.md)** method to do the conversion.


## Example

These two examples set the bottom margin of Sheet1 to 0.5 inch (36 points).


```vb
Worksheets("Sheet1").PageSetup.BottomMargin = _ 
 Application.InchesToPoints(0.5) 
 
Worksheets("Sheet1").PageSetup.BottomMargin = 36
```

This example displays the current setting for the bottom margin on Sheet1.




```
marginInches = Worksheets("Sheet1").PageSetup.BottomMargin / _ 
 Application.InchesToPoints(1) 
MsgBox "The current bottom margin is " &; marginInches &; " inches"
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-excel.md)

