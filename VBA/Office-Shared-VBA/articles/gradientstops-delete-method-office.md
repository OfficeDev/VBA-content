---
title: GradientStops.Delete Method (Office)
ms.prod: office
api_name:
- Office.GradientStops.Delete
ms.assetid: 3f31656a-498d-57d1-1464-b2439718ef89
ms.date: 06/08/2017
---


# GradientStops.Delete Method (Office)

Removes a gradient stop.


## Syntax

 _expression_. **Delete**( **_Index_** )

 _expression_ An expression that returns a **GradientStops** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional|**Integer**|The index number of the gradient stop.|

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


## See also


#### Concepts


[GradientStops Object](gradientstops-object-office.md)
#### Other resources


[GradientStops Object Members](gradientstops-members-office.md)

