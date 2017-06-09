---
title: ParagraphFormat.LineRuleWithin Property (PowerPoint)
keywords: vbapp10.chm576007
f1_keywords:
- vbapp10.chm576007
ms.prod: powerpoint
api_name:
- PowerPoint.ParagraphFormat.LineRuleWithin
ms.assetid: 0bf91b11-fe28-eec8-75f8-8fccbed19f5c
ms.date: 06/08/2017
---


# ParagraphFormat.LineRuleWithin Property (PowerPoint)

Determines whether line spacing between base lines is set to a specific number of points or lines. Read/write.


## Syntax

 _expression_. **LineRuleWithin**

 _expression_ A variable that represents a **ParagraphFormat** object.


### Return Value

MsoTriState


## Remarks

The value of the  **LineRuleWithin** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|Line spacing between base lines is set to a specific number of points.|
|**msoTrue**| Line spacing between base lines is set to a specific number of lines.|

## Example

This example displays a message box that shows the setting for line spacing for text in shape two on slide one in the active presentation.


```vb
With Application.ActivePresentation.Slides(1).Shapes(2).TextFrame

    With .TextRange.ParagraphFormat

        ls = .SpaceWithin

        If .LineRuleWithin Then

            lsUnits = " lines"

        Else

            lsUnits = " points"

        End If

    End With

End With

MsgBox "Current line spacing: " &; ls &; lsUnits
```


## See also


#### Concepts


[ParagraphFormat Object](paragraphformat-object-powerpoint.md)

