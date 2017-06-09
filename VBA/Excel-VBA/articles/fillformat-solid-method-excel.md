---
title: FillFormat.Solid Method (Excel)
keywords: vbaxl10.chm115007
f1_keywords:
- vbaxl10.chm115007
ms.prod: excel
api_name:
- Excel.FillFormat.Solid
ms.assetid: 5db7e000-7449-6bbc-192f-8b718ccffac6
ms.date: 06/08/2017
---


# FillFormat.Solid Method (Excel)

Sets the specified fill to a uniform color. Use this method to convert a gradient, textured, patterned, or background fill back to a solid fill.


## Syntax

 _expression_ . **Solid**

 _expression_ A variable that represents a **FillFormat** object.


## Example

This example converts all fills on  `myDocument` to uniform red fills.


```vb
Set myDocument = Worksheets(1) 
For Each s In myDocument.Shapes 
 With s.Fill 
 .Solid 
 .ForeColor.RGB = RGB(255, 0, 0) 
 End With 
Next
```


## See also


#### Concepts


[FillFormat Object](fillformat-object-excel.md)

