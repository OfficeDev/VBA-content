---
title: FillFormat.GradientStops Property (PowerPoint)
keywords: vbapp10.chm552025
f1_keywords:
- vbapp10.chm552025
ms.prod: powerpoint
api_name:
- PowerPoint.FillFormat.GradientStops
ms.assetid: dd0c2c5a-81f1-b008-5b2f-5248241ac0db
ms.date: 06/08/2017
---


# FillFormat.GradientStops Property (PowerPoint)

 Returns the **[GradientStops](http://msdn.microsoft.com/library/365949f0-29b3-76e1-1163-2ac870f68f7a%28Office.15%29.aspx)** collection associated with the specified fill format. Read-only.


## Syntax

 _expression_. **GradientStops**

 _expression_ An expression that returns a **FillFormat** object.


### Return Value

GradientStops


## Remarks

You can use the  **[GradientStops.Insert](http://msdn.microsoft.com/library/98aec7ed-44f9-c9b4-7a1a-e5b9a1d26d95%28Office.15%29.aspx)** method to add gradient stops to the **GradientStops** collection for the specified object.


## Example

The following example shows how to add a gradient stop at the 50% position to the  **GradientStops** collection of the fill format of the first shape on the first slide of the active presentation. For this example to work, the shape must already have a gradient fill applied.


```vb
Public Sub GradientStops_Example() 
 
    Dim pptShape As Shape 
    Dim pptFillFormat As FillFormat 
    Dim pptGradientStops As GradientStops 
     
    Set pptShape = ActivePresentation.Slides(1).Shapes(1) 
    Set pptFillFormat = pptShape.Fill 
    Set pptGradientStops = pptFillFormat.GradientStops 
     
    pptGradientStops.Insert RGB(255, 0, 255), 0.5 
     
End Sub
```


## See also


#### Concepts


[FillFormat Object](fillformat-object-powerpoint.md)

