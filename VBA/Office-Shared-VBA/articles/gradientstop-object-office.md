---
title: GradientStop Object (Office)
ms.prod: office
api_name:
- Office.GradientStop
ms.assetid: b5003bfc-9ac6-fd56-f214-a0d99db0cf07
ms.date: 06/08/2017
---


# GradientStop Object (Office)

Represents one gradient stop.


## Remarks

Gradients are a smooth transition from one color state to another. The endpoints of these sections are called stops.


## Example

The following example adds three gradient color stops and then deletes the first gradient stop.


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
 GradientShapeFill.GradientStops.Delete (1) 
End Sub
```


## Properties



|**Name**|
|:-----|
|[Application](gradientstop-application-property-office.md)|
|[Color](gradientstop-color-property-office.md)|
|[Creator](gradientstop-creator-property-office.md)|
|[Position](gradientstop-position-property-office.md)|
|[Transparency](gradientstop-transparency-property-office.md)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
