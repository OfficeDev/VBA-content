---
title: TextFrame.HasText Property (PowerPoint)
keywords: vbapp10.chm558007
f1_keywords:
- vbapp10.chm558007
ms.prod: powerpoint
api_name:
- PowerPoint.TextFrame.HasText
ms.assetid: 7bce3bae-38e7-d9d4-b67c-9454fafc620f
ms.date: 06/08/2017
---


# TextFrame.HasText Property (PowerPoint)

Returns whether the specified shape has text associated with it. Read-only.


## Syntax

 _expression_. **HasText**

 _expression_ A variable that represents a **TextFrame** object.


### Return Value

MsoTriState


## Remarks

The value of the  **HasText** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The specified shape does not have text associated with it. |
|**msoTrue**| The specified shape has text associated with it.|

## Example

If shape two on  `myDocument` contains text, this example resizes the shape to fit the text.


```vb
Set myDocument = ActivePresentation.Slides(1)

With myDocument.Shapes(2).TextFrame

    If .HasText Then .AutoSize = ppAutoSizeShapeToFitText

End With
```


## See also


#### Concepts


[TextFrame Object](textframe-object-powerpoint.md)

