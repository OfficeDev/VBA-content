---
title: GradientStop Object (Office)
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.GradientStop
ms.assetid: b5003bfc-9ac6-fd56-f214-a0d99db0cf07
---


# GradientStop Object (Office)

Represents one gradient stop.


## Remarks

Gradients are a smooth transition from one color state to another. The endpoints of these sections are called stops.


## Example

The following example adds three gradient color stops and then deletes the first gradient stop.


```vb
Sub gradients() 
 Set myDocument = ActivePresentation.Slides(1) 
 Set GradientShapeFill = myDocument.Shapes.AddShape(msoShapeRectangle, 90, 90, 90, 80).Fill 
 With GradientShapeFill 
 .ForeColor.RGB = RGB(0, 128, 128) 
 .OneColorGradient msoGradientHorizontal, 1, 1 
 .GradientStops.Insert RGB(255, 0, 0), 0.25 
 .GradientStops.Insert RGB(0, 255, 0), 0.5 
 .GradientStops.Insert RGB(0, 0, 255), 0.75 
 End With 
 GradientShapeFill.GradientStops.Delete (1) 
End Sub
```


## See also


#### Concepts


[Object Model Reference](../../Office-Shared-VBA/articles/reference-object-library-reference-for-office.md)

