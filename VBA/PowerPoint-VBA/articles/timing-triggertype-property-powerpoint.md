---
title: Timing.TriggerType Property (PowerPoint)
keywords: vbapp10.chm653004
f1_keywords:
- vbapp10.chm653004
ms.prod: powerpoint
api_name:
- PowerPoint.Timing.TriggerType
ms.assetid: a868c747-6088-df48-3c93-50f4ab24ae85
ms.date: 06/08/2017
---


# Timing.TriggerType Property (PowerPoint)

Represents the trigger that starts an animation. Read/write.


## Syntax

 _expression_. **TriggerType**

 _expression_ A variable that represents a **Timing** object.


### Return Value

MsoAnimTriggerType


## Remarks

The value of the  **TriggerType** property can be one of these **MsoAnimTriggerType** constants. The default is **msoAnimTriggerOnPageClick**.


||
|:-----|
|**msoAnimTriggerAfterPrevious**|
|**msoAnimTriggerMixed**|
|**msoAnimTriggerNone**|
|**msoAnimTriggerOnPageClick**|
|**msoAnimTriggerOnShapeClick**|
|**msoAnimTriggerWithPrevious**|

## Example

The following example adds a shape to a slide, adds an animation to the shape, and instructs the shape to begin the animation three seconds after it is clicked.


```vb
Sub AddShapeSetTiming() 
 
    Dim effDiamond As Effect 
    Dim shpRectangle As Shape 
 
    Set shpRectangle = ActivePresentation.Slides(1).Shapes _ 
        .AddShape(Type:=msoShapeRectangle, Left:=100, _ 
        Top:=100, Width:=50, Height:=50) 
    Set effDiamond = ActivePresentation.Slides(1).TimeLine.MainSequence _ 
        .AddEffect(Shape:=shpRectangle, effectId:=msoAnimEffectPathDiamond) 
 
    With effDiamond.Timing 
        .Duration = 5 
        .TriggerType = msoAnimTriggerWithPrevious
        .TriggerDelayTime = 3 
    End With 
 
End Sub
```


## See also


#### Concepts


[Timing Object](timing-object-powerpoint.md)

