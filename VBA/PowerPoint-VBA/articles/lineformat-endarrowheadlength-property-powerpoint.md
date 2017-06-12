---
title: LineFormat.EndArrowheadLength Property (PowerPoint)
keywords: vbapp10.chm553007
f1_keywords:
- vbapp10.chm553007
ms.prod: powerpoint
api_name:
- PowerPoint.LineFormat.EndArrowheadLength
ms.assetid: e7e183f6-fc85-0a5f-c1c1-f182c8020c20
ms.date: 06/08/2017
---


# LineFormat.EndArrowheadLength Property (PowerPoint)

Returns or sets the length of the arrowhead at the end of the specified line. Read/write.


## Syntax

 _expression_. **EndArrowheadLength**

 _expression_ A variable that represents an **LineFormat** object.


### Return Value

MsoArrowheadLength


## Remarks

The  **EndArrowheadLength** property value can be one of these **MsoArrowheadLength** constants.


||
|:-----|
|**msoArrowheadLengthMedium**|
|**msoArrowheadLengthMixed**|
|**msoArrowheadLong**|
|**msoArrowheadShort**|

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

