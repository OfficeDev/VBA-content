---
title: Font.Superscript Property (Excel)
keywords: vbaxl10.chm559085
f1_keywords:
- vbaxl10.chm559085
ms.prod: excel
api_name:
- Excel.Font.Superscript
ms.assetid: 23a5d707-d92a-6591-beaf-8fc62f4d3237
ms.date: 06/08/2017
---


# Font.Superscript Property (Excel)

 **True** if the font is formatted as superscript; **False** by default. Read/write **Variant** .


## Syntax

 _expression_ . **Superscript**

 _expression_ A variable that represents a **Font** object.


## Example

This example makes the last character in cell A1 a superscript character.


```vb
n = Worksheets("Sheet1").Range("A1").Characters.Count 
Worksheets("Sheet1").Range("A1") _ 
 .Characters(n, 1).Font.Superscript = True
```


## See also


#### Concepts


[Font Object](font-object-excel.md)

