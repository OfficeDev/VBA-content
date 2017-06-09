---
title: Font.Underline Property (PowerPoint)
keywords: vbapp10.chm575008
f1_keywords:
- vbapp10.chm575008
ms.prod: powerpoint
api_name:
- PowerPoint.Font.Underline
ms.assetid: ee21ab18-b131-7e4d-de19-93c9b7549d3b
ms.date: 06/08/2017
---


# Font.Underline Property (PowerPoint)

Determines whether the specified text (for the  **Font** object) or the font style (for the **FontInfo** object) is underlined. Read/write.


## Syntax

 _expression_. **Underline**

 _expression_ A variable that represents an **Font** object.


### Return Value

MsoTriState


## Remarks

The value of the  **Underline** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The specified text (or font style) isn't underlined.|
|**msoTriStateMixed**|Some characters are underlined (for the specified text) and some aren't. |
|**msoTrue**| The specified text (or font style) is underlined.|

## Example

This example sets the formatting for the text in shape two on slide one in the active presentation.


```vb
With Application.ActivePresentation.Slides(1).Shapes(2)

    With .TextFrame.TextRange.Font

        .Size = 32

        .Name = "Palatino"

        .Underline = msoTrue

    End With

End With
```


## See also


#### Concepts


[Font Object](font-object-powerpoint.md)

