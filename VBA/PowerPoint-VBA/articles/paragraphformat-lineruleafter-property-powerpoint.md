---
title: ParagraphFormat.LineRuleAfter Property (PowerPoint)
keywords: vbapp10.chm576006
f1_keywords:
- vbapp10.chm576006
ms.prod: powerpoint
api_name:
- PowerPoint.ParagraphFormat.LineRuleAfter
ms.assetid: fd206688-2217-303d-bb7e-fa3b00b0f188
ms.date: 06/08/2017
---


# ParagraphFormat.LineRuleAfter Property (PowerPoint)

Determines whether line spacing after the last line in each paragraph is set to a specific number of points or lines. Read/write.


## Syntax

 _expression_. **LineRuleAfter**

 _expression_ A variable that represents a **ParagraphFormat** object.


### Return Value

MsoTriState


## Remarks

The value of the  **LineRuleAfter** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|Line spacing after the last line in each paragraph is set to a specific number of points.|
|**msoTrue**| Line spacing after the last line in each paragraph is set to a specific number of lines.|

## Example

This example displays a message box that shows the setting for space after paragraphs for the text in shape two on slide one in the active presentation.


```vb
With Application.ActivePresentation.Slides(1).Shapes(2).TextFrame

    With .TextRange.ParagraphFormat

        sa = .SpaceAfter

        If .LineRuleAfter Then

            saUnits = " lines"

        Else

            saUnits = " points"

        End If

    End With

End With

MsgBox "Current spacing after paragraphs: " &; sa &; saUnits
```


## See also


#### Concepts


[ParagraphFormat Object](paragraphformat-object-powerpoint.md)

