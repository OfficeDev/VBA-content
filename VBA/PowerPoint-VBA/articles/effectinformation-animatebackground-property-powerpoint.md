---
title: EffectInformation.AnimateBackground Property (PowerPoint)
keywords: vbapp10.chm655004
f1_keywords:
- vbapp10.chm655004
ms.prod: powerpoint
api_name:
- PowerPoint.EffectInformation.AnimateBackground
ms.assetid: 37e9bfb5-3661-a3eb-d148-90d504f0e450
ms.date: 06/08/2017
---


# EffectInformation.AnimateBackground Property (PowerPoint)

Returns  **msoTrue** if the specified effect is a background animation. Read-only.


## Syntax

 _expression_. **AnimateBackground**

 _expression_ A variable that represents an **EffectInformation** object.


## Remarks

Use the [TextLevelEffect](animationsettings-textleveleffect-property-powerpoint.md)and  **[TextUnitEffect](effectinformation-textuniteffect-property-powerpoint.md)** properties to control the animation of text attached to the specified shape.

If this property is set to  **msoTrue** and the **TextLevelEffect** property is set to **ppAnimateByAllLevels**, the shape and its text are animated simultaneously. If this property is set to **msoTrue** and the **TextLevelEffect** property is set to anything other than **ppAnimateByAllLevels**, the shape is animated immediately before the text is animated.

You won't see effects of setting this property unless the specified shape is animated. For a shape to be animated, the  **TextLevelEffect** property for the shape must be set to something other than **ppAnimateLevelNone**, and either the **[Animate](animationsettings-animate-property-powerpoint.md)** property must be set to **msoTrue**, or the **[EntryEffect](animationsettings-entryeffect-property-powerpoint.md)** property must be set to a constant other than **ppEffectNone**.

The value returned by the  **AnimateBackground** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The specified effect is not a background animation.|
|**msoTrue**| The specified effect is a background animation.|

## Example

This example changes the direction of the animation if the background is currently animated.


```vb
Sub ChangeAnimationDirection()

    With ActivePresentation.Slides(1).TimeLine.MainSequence(1)

        If .EffectInformation.AnimateBackground = msoTrue Then

            .EffectParameters.Direction = msoAnimDirectionTopLeft

        End If

    End With

End Sub
```


## See also


#### Concepts


[EffectInformation Object](effectinformation-object-powerpoint.md)


