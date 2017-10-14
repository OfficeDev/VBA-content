---
title: PageSetup.LeftMargin Property (Excel)
keywords: vbaxl10.chm473088
f1_keywords:
- vbaxl10.chm473088
ms.prod: excel
api_name:
- Excel.PageSetup.LeftMargin
ms.assetid: 5d52ca64-6fe7-5c0e-63ab-036aa5119bb2
ms.date: 06/08/2017
---


# PageSetup.LeftMargin Property (Excel)

Returns or sets the size of the left margin, in points. Read/write  **Double** .


## Syntax

 _expression_ . **LeftMargin**

 _expression_ A variable that represents a **PageSetup** object.


## Remarks

Margins are set or returned in points. Use the  **InchesToPoints** method or the **CentimetersToPoints** method to convert measurements from inches or centimeters.


## Example

This example sets the left margin of Sheet1 to 1.5 inches.


```vb
Worksheets("Sheet1").PageSetup.LeftMargin = _ 
 Application.InchesToPoints(1.5)
```

This example sets the left margin of Sheet1 to 2 centimeters.




```vb
Worksheets("Sheet1").PageSetup.LeftMargin = _ 
 Application.CentimetersToPoints(2)
```

This example displays the current left-margin setting for Sheet1.




```
marginInches = Worksheets("Sheet1").PageSetup.LeftMargin / _ 
 Application.InchesToPoints(1) 
MsgBox "The current left margin is " &; marginInches &; " inches"
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-excel.md)

