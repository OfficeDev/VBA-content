---
title: EffectInformation.Dim Property (PowerPoint)
keywords: vbapp10.chm655007
f1_keywords:
- vbapp10.chm655007
ms.prod: powerpoint
api_name:
- PowerPoint.EffectInformation.Dim
ms.assetid: c2ffb40a-01e9-a77f-77dc-34262ed064cd
ms.date: 06/08/2017
---


# EffectInformation.Dim Property (PowerPoint)

Returns a  **[ColorFormat](colorformat-object-powerpoint.md)** object that represents the color to dim to after an animation is finished.


## Syntax

 _expression_. **Dim**

 _expression_ A variable that represents an **EffectInformation** object.


### Return Value

ColorFormat


## Example

This example displays the color to dim to after the animation.


```vb
Sub QueryDimColor()

   Dim effDim As Effect

   Set effDim = ActivePresentation.Slides(1).TimeLine.MainSequence(1)

   MsgBox effDim.EffectInformation.Dim

End Sub
```


## See also


#### Concepts


[EffectInformation Object](effectinformation-object-powerpoint.md)


