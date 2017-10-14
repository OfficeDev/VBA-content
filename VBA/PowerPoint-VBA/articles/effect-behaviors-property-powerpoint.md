---
title: Effect.Behaviors Property (PowerPoint)
keywords: vbapp10.chm652017
f1_keywords:
- vbapp10.chm652017
ms.prod: powerpoint
api_name:
- PowerPoint.Effect.Behaviors
ms.assetid: e5335758-2f92-ccbc-a665-b6d5947e79f2
ms.date: 06/08/2017
---


# Effect.Behaviors Property (PowerPoint)

Returns a specified slide animation behavior as an  **[AnimationBehaviors](animationbehaviors-object-powerpoint.md)** collection.


## Syntax

 _expression_. **Behaviors**

 _expression_ A variable that represents an **Effect** object.


### Return Value

AnimationBehaviors


## Remarks

To return a single  **[AnimationBehavior](animationbehavior-object-powerpoint.md)** object in the **AnimationBehaviors** collection, use the **[Item](animationbehaviors-item-method-powerpoint.md)** method or **Behaviors** (index), where index is the index number of the **AnimationBehavior** object in the **AnimationBehaviors** collection.


## Example

The following example returns a specific animation behavior type in the active presentation.


```vb
Sub ReturnTypeValue
    MsgBox ActiveWindow.Selection.SlideRange(1).TimeLine _
        .MainSequence(1).Behaviors.Item(1).Type
End Sub
```


## See also


#### Concepts



[Effect Object](effect-object-powerpoint.md)

