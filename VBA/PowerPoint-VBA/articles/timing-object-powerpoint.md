---
title: Timing Object (PowerPoint)
keywords: vbapp10.chm653000
f1_keywords:
- vbapp10.chm653000
ms.prod: powerpoint
api_name:
- PowerPoint.Timing
ms.assetid: 11f7dab2-f9ed-1883-ab74-93f1be481af6
ms.date: 06/08/2017
---


# Timing Object (PowerPoint)

Represents timing properties for an animation effect.


## Remarks

Use the following read/write properties of the  **Timing** object to manipulate animation timing effects.



|**Use this property**|**To change this...**|
|:-----|:-----|
|[Accelerate](timing-accelerate-property-powerpoint.md)|Percentage of the duration over which acceleration should take place|
|[AutoReverse](timing-autoreverse-property-powerpoint.md)|Whether an effect should play forward and then reverse, thereby doubling the duration|
|[Decelerate](timing-decelerate-property-powerpoint.md)|Percentage of the duration over which acceleration should take place|
|[Duration](slideshowtransition-duration-property-powerpoint.md)|Length of animation (in seconds)|
|[RepeatCount](timing-repeatcount-property-powerpoint.md)|Number of times to repeat the animation|
|[RepeatDuration](timing-repeatduration-property-powerpoint.md)|How long should the repeats last (in seconds)|
|[Restart](timing-restart-property-powerpoint.md)|Restart behavior of an animation node|
|[RewindAtEnd](timing-rewindatend-property-powerpoint.md)|Whether an objects return to its beginning position after an effect has ended|
|[SmoothStart](timing-smoothstart-property-powerpoint.md)|Whether an effect accelerates when it starts|
|[SmoothEnd](timing-smoothend-property-powerpoint.md)|Whether an effect decelerates when it ends|
|[TriggerDelayTime](timing-triggerdelaytime-property-powerpoint.md)|Delay time from when the trigger is enabled (in seconds)|
|[TriggerShape](timing-triggershape-property-powerpoint.md)|Which shape is associated with the timing effect|
|[TriggerType](timing-triggertype-property-powerpoint.md)|How the timing effect is triggered|

## Example

To return a  **Timing** object, use the[Timing](animationbehavior-timing-property-powerpoint.md)property of the  **[AnimationBehavior](animationbehavior-object-powerpoint.md)** or **[Effect](effect-object-powerpoint.md)** object. The following example sets timing duration information for the main animation.


```vb
ActiveWindow.Selection.SlideRange(1).TimeLine.MainSequence(1).Timing.Duration = 5
```


## See also


#### Concepts


[PowerPoint Object Model Reference](object-model-powerpoint-vba-reference.md)

