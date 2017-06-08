---
title: TextFrame2.HasText Property (PowerPoint)
keywords: vbapp10.chm678015
f1_keywords:
- vbapp10.chm678015
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame2.HasText
ms.assetid: 50b2c7fa-49f9-6aeb-dcb0-8acaf7aefec7
ms.date: 06/08/2017
---


# TextFrame2.HasText Property (PowerPoint)

 Indicates whether the shape that contains the specified text frame has text associated with it. Read-only.


## Syntax

 _expression_. **HasText**

 _expression_ An expression that returns a **TextFrame2** object.


### Return Value

MsoTriState


## Remarks

The value of the  **HasText** property can be one of the following **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The specified text frame does not have text.|
|**msoTrue**| The specified text frame has text.|

## Example

The followin example tests whether shape two on slide one contains text, and if it does, resizes the shape to fit the text.


```vb
Public Sub HasText_Example()



    Dim pptSlide As Slide

    Set pptSlide = ActivePresentation.Slides(1)

    With pptSlide.Shapes(2).TextFrame

        If .HasText Then .AutoSize = ppAutoSizeShapeToFitText

    End With



End Sub
```


## See also


#### Concepts


[TextFrame2 Object](textframe2-object-powerpoint.md)

