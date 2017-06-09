---
title: SlideShowTransition.AdvanceOnClick Property (PowerPoint)
keywords: vbapp10.chm539003
f1_keywords:
- vbapp10.chm539003
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowTransition.AdvanceOnClick
ms.assetid: 0f517795-ea23-4c94-fad9-cc2e6c1cd5e6
ms.date: 06/08/2017
---


# SlideShowTransition.AdvanceOnClick Property (PowerPoint)

Determines whether the specified slide advances when it is clicked during a slide show. Read/write.


## Syntax

 _expression_. **AdvanceOnClick**

 _expression_ A variable that represents an **SlideShowTransition** object.


### Return Value

MsoTriState


## Remarks

To set the slide to advance automatically after a certain amount of time elapses, set the  **[AdvanceOnTime](slideshowtransition-advanceontime-property-powerpoint.md)** property to **True** and set the **[AdvanceTime](slideshowtransition-advancetime-property-powerpoint.md)** property to the amount of time you want the slide to be shown. If you set both the **AdvanceOnClick** and the **AdvanceOnTime** properties to **True**, the slide advances either when it is clicked or when the specified amount of time has elapsed ? whichever comes first.

The value of the  **AdvanceOnClick** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The specified slide does not advance when it is clicked during a slide show.|
|**msoTrue**| The specified slide advances when it is clicked during a slide show.|

## Example

This example sets slide one in the active presentation to advance after five seconds have passed or when the mouse is clicked ? whichever occurs first.


```vb
With ActivePresentation.Slides(1).SlideShowTransition

    .AdvanceOnClick = msoTrue

    .AdvanceOnTime = msoTrue

    .AdvanceTime = 5

End With
```


## See also


#### Concepts


[SlideShowTransition Object](slideshowtransition-object-powerpoint.md)

