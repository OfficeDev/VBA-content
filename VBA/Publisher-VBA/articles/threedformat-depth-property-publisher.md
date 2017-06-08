---
title: ThreeDFormat.Depth Property (Publisher)
keywords: vbapb10.chm3801344
f1_keywords:
- vbapb10.chm3801344
ms.prod: publisher
api_name:
- Publisher.ThreeDFormat.Depth
ms.assetid: b6b46ddb-e3dd-0f9a-1a67-6433bb9ea89a
ms.date: 06/08/2017
---


# ThreeDFormat.Depth Property (Publisher)

Returns or sets a  **Variant** indicating the depth of the shape's extrusion. Read/write.


## Syntax

 _expression_. **Depth**

 _expression_A variable that represents a  **ThreeDFormat** object.


### Return Value

Variant


## Remarks

Numeric values are evaluated in points; strings can be in any units supported by Microsoft Publisher (for example, "2.5 in").

Positive values produce an extrusion whose front face is the original shape; negative values produce an extrusion whose back face is the original shape. The valid range is -600 through 9600 points, or the equivalent distance in all other units.


## Example

This example adds an oval to the active publication, and then specifies that the oval be extruded to a depth of 50 points and that the extrusion be purple.


```vb
Dim shpNew As Shape 
 
Set shpNew = ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShapeOval, _ 
 Left:=90, Top:=90, Width:=90, Height:=40) 
 
With shpNew.ThreeD 
 .Visible = True 
 .Depth = 50 
 .ExtrusionColor.RGB = RGB(255, 100, 255) 
End With 

```


