---
title: Font.Embeddable Property (PowerPoint)
keywords: vbapp10.chm575013
f1_keywords:
- vbapp10.chm575013
ms.prod: powerpoint
api_name:
- PowerPoint.Font.Embeddable
ms.assetid: 50824587-0371-e7eb-8885-370f97b8bf0c
ms.date: 06/08/2017
---


# Font.Embeddable Property (PowerPoint)

Determines whether the specified font can be embedded in the presentation. Read-only.


## Syntax

 _expression_. **Embeddable**

 _expression_ A variable that represents an **Font** object.


### Return Value

MsoTriState


## Remarks

The value of the  **Embeddable** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The specified font cannot be embedded in the presentation.|
|**msoTrue**| The specified font can be embedded in the presentation.|

## Example

This example checks each font used in the active presentation to determine whether it is embeddable in the presentation.


```vb
For Each usedFont In Presentations(1).Fonts

    If usedFont.Embeddable Then

        MsgBox usedFont.Name &; ": Embeddable"

    Else

        MsgBox usedFont.Name &; ": Not embeddable"

    End If

Next usedFont
```


## See also


#### Concepts


[Font Object](font-object-powerpoint.md)

