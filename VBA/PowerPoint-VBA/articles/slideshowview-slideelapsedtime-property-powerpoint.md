---
title: SlideShowView.SlideElapsedTime Property (PowerPoint)
keywords: vbapp10.chm513009
f1_keywords:
- vbapp10.chm513009
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowView.SlideElapsedTime
ms.assetid: e9250ea3-c37e-ebed-c8a8-9774dab77f37
ms.date: 06/08/2017
---


# SlideShowView.SlideElapsedTime Property (PowerPoint)

Returns the number of seconds that the current slide has been displayed. Read/write.


## Syntax

 _expression_. **SlideElapsedTime**

 _expression_ A variable that represents a **SlideShowView** object.


### Return Value

Long


## Remarks

Use the  **[ResetSlideTime](slideshowview-resetslidetime-method-powerpoint.md)** method to reset the elapsed time for the slide that's currently displayed.


## Example

This example sets a variable to the elapsed time for the slide that's currently displayed in slide show window one and then displays the value of the variable.


```
currTime = SlideShowWindows(1).View.SlideElapsedTime

MsgBox currTime
```


## See also


#### Concepts


[SlideShowView Object](slideshowview-object-powerpoint.md)

