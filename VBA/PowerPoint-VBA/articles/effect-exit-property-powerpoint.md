---
title: Effect.Exit Property (PowerPoint)
keywords: vbapp10.chm652016
f1_keywords:
- vbapp10.chm652016
ms.prod: powerpoint
api_name:
- PowerPoint.Effect.Exit
ms.assetid: 0f4d74d4-ce88-f9b9-7de5-0e42edf12967
ms.date: 06/08/2017
---


# Effect.Exit Property (PowerPoint)

Determines whether the animation effect is an exit effect. Read/write.


## Syntax

 _expression_. **Exit**

 _expression_ A variable that represents an **Effect** object.


### Return Value

MsoTriState


## Remarks

The value of the  **Exit** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The effect is not an exit effect.|
|**msoTrue**| The effect is an exit effect.|

## Example

This example displays whether the specified animation is an exit animation effect.


```vb
Sub EffectExit()

    Dim effMain As Effect

    Set effMain = ActivePresentation.Slides(1).TimeLine.MainSequence(1)

    If effMain.Exit = msoTrue Then

        MsgBox "This is an exit animation effect."

    Else

        MsgBox "This is not an exit animation effect."

    End If

End Sub
```


## See also


#### Concepts



[Effect Object](effect-object-powerpoint.md)

