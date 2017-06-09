---
title: TextFrame.AutoSize Property (PowerPoint)
keywords: vbapp10.chm558012
f1_keywords:
- vbapp10.chm558012
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame.AutoSize
ms.assetid: 771f5217-abce-f70d-743d-e17532dabd9e
ms.date: 06/08/2017
---


# TextFrame.AutoSize Property (PowerPoint)

Returns or sets a value that indicates whether the size of the specified shape is changed automatically to fit text within its boundaries. Read/write.


## Syntax

 _expression_. **AutoSize**

 _expression_ A variable that represents an **TextFrame** object.


### Return Value

PpAutoSize


## Remarks

The value of the  **AutoSize** property can be one of these **PpAutoSize** constants.


||
|:-----|
|**ppAutoSizeMixed**|
|**ppAutoSizeNone**|
|**ppAutoSizeShapeToFitText**|

## Example

This example adjusts the size of the title bounding box on slide one to fit the title text.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(1)

    If .TextFrame.TextRange.Characters.Count < 50 Then

        .TextFrame.AutoSize = ppAutoSizeShapeToFitText

    End If

End With
```


## See also


#### Concepts


[TextFrame Object](textframe-object-powerpoint.md)

