---
title: LineFormat.DashStyle Property (PowerPoint)
keywords: vbapp10.chm553006
f1_keywords:
- vbapp10.chm553006
ms.prod: powerpoint
api_name:
- PowerPoint.LineFormat.DashStyle
ms.assetid: 7fc898b4-1eea-21fc-52e5-0ec92bde527f
ms.date: 06/08/2017
---


# LineFormat.DashStyle Property (PowerPoint)

Returns or sets the dash style for the specified line. Read/write.


## Syntax

 _expression_. **DashStyle**

 _expression_ A variable that represents a **LineFormat** object.


### Return Value

MsoLineDashStyle


## Remarks

The value of the  **DashStyle** property can be one of these **MsoLineDashStyle** constants.


||
|:-----|
|**msoLineDash**|
|**msoLineDashDot**|
|**msoLineDashDotDot**|
|**msoLineDashStyleMixed**|
|**msoLineLongDash**|
|**msoLineLongDashDot**|
|**msoLineRoundDot**|
|**msoLineSolid**|
|**msoLineSquareDot**|

## Example

This example adds a blue dashed line to  `myDocument`.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes.AddLine(10, 10, 250, 250).Line

    .DashStyle = msoLineDashDotDot

    .ForeColor.RGB = RGB(50, 0, 128)

End With
```


## See also


#### Concepts


[LineFormat Object](lineformat-object-powerpoint.md)

