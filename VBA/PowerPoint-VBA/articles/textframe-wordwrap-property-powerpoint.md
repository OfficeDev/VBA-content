---
title: TextFrame.WordWrap Property (PowerPoint)
keywords: vbapp10.chm558013
f1_keywords:
- vbapp10.chm558013
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame.WordWrap
ms.assetid: f6077142-9afd-b274-7301-3e63d962e7b3
ms.date: 06/08/2017
---


# TextFrame.WordWrap Property (PowerPoint)

Determines whether lines break automatically to fit inside the shape. Read/write.


## Syntax

 _expression_. **WordWrap**

 _expression_ A variable that represents a **TextFrame** object.


### Return Value

MsoTriState


## Remarks

The value of the  **WordWrap** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**| Lines do not break automatically to fit inside the shape.|
|**msoTrue**| Lines break automatically to fit inside the shape.|

## Example

This example adds a rectangle that contains text to  `myDocument` and then turns off word wrapping in the new rectangle.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes.AddShape(msoShapeRectangle, _
        0, 0, 100, 300).TextFrame
    .TextRange.Text = _
        "Here is some test text that is too long for this box"
    .WordWrap = False
End With
```


## See also


#### Concepts


[TextFrame Object](textframe-object-powerpoint.md)

