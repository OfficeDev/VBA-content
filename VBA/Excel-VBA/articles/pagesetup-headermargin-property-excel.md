---
title: PageSetup.HeaderMargin Property (Excel)
keywords: vbaxl10.chm473085
f1_keywords:
- vbaxl10.chm473085
ms.prod: excel
api_name:
- Excel.PageSetup.HeaderMargin
ms.assetid: c22feaf6-c9f5-f285-a8f6-75753a1e9cff
ms.date: 06/08/2017
---


# PageSetup.HeaderMargin Property (Excel)

Returns or sets the distance from the top of the page to the header, in points. Read/write  **Double** .


## Syntax

 _expression_ . **HeaderMargin**

 _expression_ A variable that represents a **PageSetup** object.


## Remarks

Margins are set or returned in points. Use the  **InchesToPoints** method or the **CentimetersToPoints** method to convert measurements from inches or centimeters.


## Example

This example sets the header margin of Sheet1 to 0.5 inch.


```vb
Worksheets("Sheet1").PageSetup.HeaderMargin = _ 
 Application.InchesToPoints(0.5)
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-excel.md)

