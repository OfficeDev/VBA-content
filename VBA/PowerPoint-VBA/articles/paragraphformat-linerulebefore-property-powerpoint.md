---
title: ParagraphFormat.LineRuleBefore Property (PowerPoint)
keywords: vbapp10.chm576005
f1_keywords:
- vbapp10.chm576005
ms.prod: powerpoint
api_name:
- PowerPoint.ParagraphFormat.LineRuleBefore
ms.assetid: 2316216e-9f56-07e6-1b32-10b37a6fdc9d
ms.date: 06/08/2017
---


# ParagraphFormat.LineRuleBefore Property (PowerPoint)

Determines whether line spacing before the first line in each paragraph is set to a specific number of points or lines. Read/write.


## Syntax

 _expression_. **LineRuleBefore**

 _expression_ A variable that represents a **ParagraphFormat** object.


### Return Value

MsoTriState


## Remarks

The value of the  **LineRuleBefore** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|Line spacing before the first line in each paragraph is set to a specific number of points. |
|**msoTrue**| Line spacing before the first line in each paragraph is set to a specific number of lines.|

## Example

This example displays a message box that shows the setting for space before paragraphs for text in shape two on slide one in the active presentation.


```vb
With Application.ActivePresentation.Slides(1).Shapes(2).TextFrame

    With .TextRange.ParagraphFormat

        sb = .SpaceBefore

        If .LineRuleBefore Then

            sbUnits = " lines"

        Else

            sbUnits = " points"

        End If

    End With

End With

MsgBox "Current spacing before paragraphs: " &; sb &; sbUnits
```


## See also


#### Concepts


[ParagraphFormat Object](paragraphformat-object-powerpoint.md)

