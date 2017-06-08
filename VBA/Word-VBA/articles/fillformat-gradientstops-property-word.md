---
title: FillFormat.GradientStops Property (Word)
keywords: vbawd10.chm164102258
f1_keywords:
- vbawd10.chm164102258
ms.prod: word
api_name:
- Word.FillFormat.GradientStops
ms.assetid: 3ae72292-2b7b-69af-35d4-5f887ce3c7ce
ms.date: 06/08/2017
---


# FillFormat.GradientStops Property (Word)

Returns the [GradientStops](http://msdn.microsoft.com/library/365949f0-29b3-76e1-1163-2ac870f68f7a%28Office.15%29.aspx) collection associated with the specified fill format. Read-only.


## Syntax

 _expression_ . **GradientStops**

 _expression_ An expression that returns a **FillFormat** object.


## Remarks

Gradients are a smooth transition from one color state to another. The endpoints of these sections are called stops. You can use the [GradientStops.Insert](http://msdn.microsoft.com/library/98aec7ed-44f9-c9b4-7a1a-e5b9a1d26d95%28Office.15%29.aspx) method to add gradient stops to the **GradientStops** collection for the specified object.


## Example

The following code example shows how to add a gradient stop at the 50% position to the  **GradientStops** collection of the fill format of the first shape in the active document. For this code example to work, the shape must already have a gradient fill applied.


```vb
Public Sub GradientStops_Example() 
 
 Dim wdShape As Shape 
 Dim wdFillFormat As FillFormat 
 Dim wdGradientStops As GradientStops 
 
 Set wdShape = ActiveDocument.Shapes(1) 
 Set wdFillFormat = wdShape.Fill 
 Set wdGradientStops = wdFillFormat.GradientStops 
 
 wdGradientStops.Insert RGB(255, 0, 255), 0.5 
End Sub
```


## See also


#### Concepts


[FillFormat Object](fillformat-object-word.md)

