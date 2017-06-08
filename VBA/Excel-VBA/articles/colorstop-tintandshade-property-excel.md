---
title: ColorStop.TintAndShade Property (Excel)
keywords: vbaxl10.chm851076
f1_keywords:
- vbaxl10.chm851076
ms.prod: excel
api_name:
- Excel.ColorStop.TintAndShade
ms.assetid: 64602eee-9196-fa9b-9a09-e11a4433b4f3
ms.date: 06/08/2017
---


# ColorStop.TintAndShade Property (Excel)

Returns or sets the tint and shade of the represented object. Read/write


## Syntax

 _expression_ . **TintAndShade**

 _expression_ A variable that represents a **ColorStop** object.


### Return Value

Variant


## Example

Applies tint and shade to the active selection.


```vb
Range("A1:A10").Select 
With Selection.Interior.Gradient.ColorStop.Add(1) 
 .ThemeColor = xlThemeColorAccent1 
 .TintAndShade = 0 
End With
```


## See also


#### Concepts


[ColorStop Object](colorstop-object-excel.md)

