---
title: SlideShowView.PointerType Property (PowerPoint)
keywords: vbapp10.chm513005
f1_keywords:
- vbapp10.chm513005
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowView.PointerType
ms.assetid: 58f40da1-ae25-4604-86bc-6fb884b8fd16
ms.date: 06/08/2017
---


# SlideShowView.PointerType Property (PowerPoint)

Returns or sets the type of pointer used in the slide show. Read/write.


## Syntax

 _expression_. **PointerType**

 _expression_ A variable that represents a **SlideShowView** object.


### Return Value

PpSlideShowPointerType


## Remarks

The value of the  **PointerType** property can be one of these **PpSlideShowPointerType** constants.


||
|:-----|
|**ppSlideShowPointerAlwaysHidden**|
|**ppSlideShowPointerArrow**|
|**ppSlideShowPointerAutoArrow**|
|**ppSlideShowPointerNone**|
|**ppSlideShowPointerPen**|

## Example

This example runs a slide show of the active presentation, changes the pointer to a pen, and sets the pen color for this slide show to red.


```vb
Set currView = ActivePresentation.SlideShowSettings.Run.View

With currView

    .PointerColor.RGB = RGB(255, 0, 0)

    .PointerType = ppSlideShowPointerPen

End With
```


## See also


#### Concepts


[SlideShowView Object](slideshowview-object-powerpoint.md)

