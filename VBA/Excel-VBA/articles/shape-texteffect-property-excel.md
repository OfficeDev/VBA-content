---
title: Shape.TextEffect Property (Excel)
keywords: vbaxl10.chm636108
f1_keywords:
- vbaxl10.chm636108
ms.prod: excel
api_name:
- Excel.Shape.TextEffect
ms.assetid: 4e2920c3-340c-c113-2667-4d4779cfb59f
ms.date: 06/08/2017
---


# Shape.TextEffect Property (Excel)

Returns a  **[TextEffectFormat](texteffectformat-object-excel.md)** object that contains text-effect formatting properties for the specified shape. Read-only.


## Syntax

 _expression_ . **TextEffect**

 _expression_ A variable that represents a **Shape** object.


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


[Shape Object](shape-object-excel.md)

