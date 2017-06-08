---
title: ThreeDFormat.ExtrusionColorType Property (PowerPoint)
keywords: vbapp10.chm557009
f1_keywords:
- vbapp10.chm557009
ms.prod: powerpoint
api_name:
- PowerPoint.ThreeDFormat.ExtrusionColorType
ms.assetid: 2e6acc19-fcdf-70e2-6ddd-7142e904d225
ms.date: 06/08/2017
---


# ThreeDFormat.ExtrusionColorType Property (PowerPoint)

Returns or sets a value that indicates whether the extrusion color is based on the extruded shape's fill (the front face of the extrusion) and automatically changes when the shape's fill changes, or whether the extrusion color is independent of the shape's fill. Read/write.


## Syntax

 _expression_. **ExtrusionColorType**

 _expression_ A variable that represents an **ThreeDFormat** object.


### Return Value

MsoExtrusionColorType


## Remarks

The value of the  **ExtrusionColorType** property can be one of these **MsoExtrusionColorType** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoExtrusionColorAutomatic**|Extrusion color is based on shape fill.|
|**msoExtrusionColorCustom**| Extrusion color is independent of shape fill.|
|**msoExtrusionColorTypeMixed**|Extrusion color is partially independent of shape fill.|

## Example

If shape one on  `myDocument` has an automatic extrusion color, this example gives the extrusion a custom yellow color.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(1).ThreeD

    If .ExtrusionColorType = msoExtrusionColorAutomatic Then

        .ExtrusionColor.RGB = RGB(240, 235, 16)

    End If

End With
```


## See also


#### Concepts


[ThreeDFormat Object](threedformat-object-powerpoint.md)

