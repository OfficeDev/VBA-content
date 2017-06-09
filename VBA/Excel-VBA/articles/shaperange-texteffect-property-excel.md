---
title: ShapeRange.TextEffect Property (Excel)
keywords: vbaxl10.chm640115
f1_keywords:
- vbaxl10.chm640115
ms.prod: excel
api_name:
- Excel.ShapeRange.TextEffect
ms.assetid: 95c2ab5d-061e-f50e-fc2b-7c44ffca7ce9
ms.date: 06/08/2017
---


# ShapeRange.TextEffect Property (Excel)

Returns a  **[TextEffectFormat](texteffectformat-object-excel.md)** object that contains text-effect formatting properties for the specified shape. Read-only.


## Syntax

 _expression_ . **TextEffect**

 _expression_ A variable that represents a **ShapeRange** object.


## Example

This example sets the font style to bold for shape three on  `myDocument` if the shape is WordArt.


```vb
Set myDocument = Worksheets(1) 
With myDocument.Shapes(3) 
 If .Type = msoTextEffect Then 
 .TextEffect.FontBold = True 
 End If 
End With
```


## See also


#### Concepts


[ShapeRange Object](shaperange-object-excel.md)

