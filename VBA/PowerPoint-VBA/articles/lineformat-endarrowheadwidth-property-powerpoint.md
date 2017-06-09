---
title: LineFormat.EndArrowheadWidth Property (PowerPoint)
keywords: vbapp10.chm553009
f1_keywords:
- vbapp10.chm553009
ms.prod: powerpoint
api_name:
- PowerPoint.LineFormat.EndArrowheadWidth
ms.assetid: 5830e4ff-c630-198a-ea2b-b5d1397ea846
ms.date: 06/08/2017
---


# LineFormat.EndArrowheadWidth Property (PowerPoint)

Returns or sets the width of the arrowhead at the end of the specified line. Read/write.


## Syntax

 _expression_. **EndArrowheadWidth**

 _expression_ A variable that represents an **LineFormat** object.


### Return Value

MsoArrowheadWidth


## Remarks

The  **EndArrowheadWidth** property value can be one of these **MsoArrowheadWidth** constants.


||
|:-----|
|**msoArrowheadNarrow**|
|**msoArrowheadWide**|
|**msoArrowheadWidthMedium**|
|**msoArrowheadWidthMixed**|

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

