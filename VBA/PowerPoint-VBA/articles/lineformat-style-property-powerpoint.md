---
title: LineFormat.Style Property (PowerPoint)
keywords: vbapp10.chm553012
f1_keywords:
- vbapp10.chm553012
ms.prod: powerpoint
api_name:
- PowerPoint.LineFormat.Style
ms.assetid: 8a9b1a85-f290-97f5-c19d-6427d1214f7b
ms.date: 06/08/2017
---


# LineFormat.Style Property (PowerPoint)

Returns or sets the line style. Read/write.


## Syntax

 _expression_. **Style**

 _expression_ A variable that represents a **LineFormat** object.


### Return Value

MsoLineStyle


## Remarks

The value of the  **Style** property can be one of these **MsoLineStyle** constants.


||
|:-----|
|**msoLineSingle**|
|**msoLineStyleMixed**|
|**msoLineThickBetweenThin**|
|**msoLineThickThin**|
|**msoLineThinThick**|
|**msoLineThinThin**|

## Example

This example adds a thick, blue, compound line to  `myDocument`. The compound line consists of a thick line with a thin line on either side of it.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes.AddLine(10, 10, 250, 250).Line

    .Style = msoLineThickBetweenThin

    .Weight = 8

    .ForeColor.RGB = RGB(0, 0, 255)

End With
```


## See also


#### Concepts


[LineFormat Object](lineformat-object-powerpoint.md)

