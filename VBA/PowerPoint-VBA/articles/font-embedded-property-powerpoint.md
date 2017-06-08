---
title: Font.Embedded Property (PowerPoint)
keywords: vbapp10.chm575012
f1_keywords:
- vbapp10.chm575012
ms.prod: powerpoint
api_name:
- PowerPoint.Font.Embedded
ms.assetid: 3fd7fe50-19a9-9944-f7c8-0ba54bc07c93
ms.date: 06/08/2017
---


# Font.Embedded Property (PowerPoint)

Determines whether the specified font is embedded in the presentation. Read-only.


## Syntax

 _expression_. **Embedded**

 _expression_ A variable that represents an **Font** object.


### Return Value

MsoTriState


## Remarks

The value of the  **Embedded** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The specified font is not embedded in the presentation. |
|**msoTrue**| The specified font is embedded in the presentation.|

## Example

This example checks each font used in the active presentation to determine whether it is embedded in the presentation.


```vb
For Each usedFont In Presentations(1).Fonts

    If usedFont.Embedded Then

        MsgBox usedFont.Name &; ": Embedded"

    Else

        MsgBox usedFont.Name &; ": Not embedded"

    End If

Next usedFont
```


## See also


#### Concepts


[Font Object](font-object-powerpoint.md)

