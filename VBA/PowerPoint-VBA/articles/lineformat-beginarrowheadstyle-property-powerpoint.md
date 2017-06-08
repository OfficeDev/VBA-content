---
title: LineFormat.BeginArrowheadStyle Property (PowerPoint)
keywords: vbapp10.chm553004
f1_keywords:
- vbapp10.chm553004
ms.prod: powerpoint
api_name:
- PowerPoint.LineFormat.BeginArrowheadStyle
ms.assetid: 04f6e7f1-c76f-b70d-5fbd-daaa907fe59d
ms.date: 06/08/2017
---


# LineFormat.BeginArrowheadStyle Property (PowerPoint)

Returns or sets the style of the arrowhead at the beginning of the specified line. Read/write.


## Syntax

 _expression_. **BeginArrowheadStyle**

 _expression_ A variable that represents a **LineFormat** object.


### Return Value

MsoArrowheadStyle


## Remarks

The value of the  **BeginArrowheadStyle** property can be one of these **MsoArrowheadStyle** constants


||
|:-----|
|**msoArrowheadDiamond**|
|**msoArrowheadNone**|
|**msoArrowheadOpen**|
|**msoArrowheadOval**|
|**msoArrowheadStealth**|
|**msoArrowheadStyleMixed**|
|**msoArrowheadTriangle**|

## Example

This example adds a line to  `myDocument`. There's a short, narrow oval on the line's starting point and a long, wide triangle on its endpoint.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes.AddLine(100, 100, 200, 300).Line

    .BeginArrowheadLength = msoArrowheadShort

    .BeginArrowheadStyle = msoArrowheadOval

    .BeginArrowheadWidth = msoArrowheadNarrow

    .EndArrowheadLength = msoArrowheadLong

    .EndArrowheadStyle = msoArrowheadTriangle

    .EndArrowheadWidth = msoArrowheadWide

End With
```


## See also


#### Concepts


[LineFormat Object](lineformat-object-powerpoint.md)

