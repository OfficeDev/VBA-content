---
title: SlideShowView.State Property (PowerPoint)
keywords: vbapp10.chm513006
f1_keywords:
- vbapp10.chm513006
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowView.State
ms.assetid: 749fe106-fed4-6ccc-f127-2e8a80196309
ms.date: 06/08/2017
---


# SlideShowView.State Property (PowerPoint)

Returns or sets the state of the slide show. Read/write.


## Syntax

 _expression_. **State**

 _expression_ A variable that represents a **SlideShowView** object.


### Return Value

PpSlideShowState


## Remarks

The value of the  **State** property can be one of these **PpSlideShowState** constants.


||
|:-----|
|**ppSlideShowBlackScreen**|
|**ppSlideShowDone**|
|**ppSlideShowPaused**|
|**ppSlideShowRunning**|
|**ppSlideShowWhiteScreen**|

## Example

This example sets the view state in slide show window one to a black screen.


```
SlideShowWindows(1).View.State = ppSlideShowBlackScreen
```


## See also


#### Concepts


[SlideShowView Object](slideshowview-object-powerpoint.md)

