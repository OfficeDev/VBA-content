---
title: SlideShowTransition.AdvanceOnTime Property (PowerPoint)
keywords: vbapp10.chm539004
f1_keywords:
- vbapp10.chm539004
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowTransition.AdvanceOnTime
ms.assetid: 934c5acc-b230-2b7b-f0f2-4647cce5b62d
ms.date: 06/08/2017
---


# SlideShowTransition.AdvanceOnTime Property (PowerPoint)

Determines whether the specified slide advances automatically after a specified amount of time has elapsed. Read/write.


## Syntax

 _expression_. **AdvanceOnTime**

 _expression_ A variable that represents an **SlideShowTransition** object.


### Return Value

MsoTriState


## Remarks

Use the  **[AdvanceTime](slideshowtransition-advancetime-property-powerpoint.md)** property to specify the number of seconds after which the slide automatically advances. Set the **[AdvanceMode](slideshowsettings-advancemode-property-powerpoint.md)** property of the **SlideShowSettings** object to **ppSlideShowUseSlideTimings** to put the slide interval settings into effect for the entire slide show.

The value of the  **AdvanceOnTime** property can be one of these **MsoTriState** constants.



|**Constant**|**Description**|
|:-----|:-----|
|**msoFalse**|The specified slide does not advance automatically after a specified amount of time has elapsed. |
|**msoTrue**| The specified slide advances automatically after a specified amount of time has elapsed.|

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

