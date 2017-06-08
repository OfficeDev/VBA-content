---
title: GradientStops Object (Office)
ms.prod: office
api_name:
- Office.GradientStops
ms.assetid: 365949f0-29b3-76e1-1163-2ac870f68f7a
ms.date: 06/08/2017
---


# GradientStops Object (Office)

Contains a collection of  **GradientStop** objects.


## Remarks

Gradients are a smooth transition from one color state to another. The endpoints of these sections are called stops.


## Example

The following example creates three color gradient stops in Microsoft PowerPoint.


```
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
End Sub
```


## See also


#### Concepts


[Object Model Reference](reference-object-library-reference-for-office.md)
#### Other resources


[GradientStops Object Members](gradientstops-members-office.md)

