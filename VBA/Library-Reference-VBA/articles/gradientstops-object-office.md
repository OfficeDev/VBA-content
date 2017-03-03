---
title: GradientStops Object (Office)
ms.prod: MULTIPLEPRODUCTS
api_name:
- Office.GradientStops
ms.assetid: 365949f0-29b3-76e1-1163-2ac870f68f7a
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


## Methods



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/3f31656a-498d-57d1-1464-b2439718ef89%28Office.15%29.aspx)|
|[Insert](http://msdn.microsoft.com/library/98aec7ed-44f9-c9b4-7a1a-e5b9a1d26d95%28Office.15%29.aspx)|
|[Insert2](http://msdn.microsoft.com/library/bd9ed41d-eaeb-d3aa-6a8a-e38e2bfb9a17%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Application](http://msdn.microsoft.com/library/f4c9ca0c-9796-8290-438f-8ce0a174cb18%28Office.15%29.aspx)|
|[Count](http://msdn.microsoft.com/library/d43892a5-8abc-38fc-efc1-311dc8125575%28Office.15%29.aspx)|
|[Creator](http://msdn.microsoft.com/library/3dc34737-a6f9-7e8a-ba69-e200f53bedc5%28Office.15%29.aspx)|
|[Item](http://msdn.microsoft.com/library/0bf0ad81-0afc-ae32-be50-e5fb772a676e%28Office.15%29.aspx)|

## See also


#### Other resources


[Object Model Reference](http://msdn.microsoft.com/library/499c789a-aba2-0fad-649a-0ea964cd3b5e%28Office.15%29.aspx)
[GradientStops Object Members](http://msdn.microsoft.com/library/9cab316d-3302-a119-b02b-54eea372acee%28Office.15%29.aspx)
