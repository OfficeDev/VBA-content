---
title: SlideShowSettings.ShowType Property (PowerPoint)
keywords: vbapp10.chm514010
f1_keywords:
- vbapp10.chm514010
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowSettings.ShowType
ms.assetid: 6537dd4c-8029-3e95-7073-7701ba12a627
ms.date: 06/08/2017
---


# SlideShowSettings.ShowType Property (PowerPoint)

Returns or sets the show type for the specified slide show. Read/write.


## Syntax

 _expression_. **ShowType**

 _expression_ A variable that represents a **SlideShowSettings** object.


### Return Value

PpSlideShowType


## Remarks

The value of the  **ShowType** property can be one of these **PpSlideShowType** constants.


||
|:-----|
|**ppShowTypeKiosk**|
|**ppShowTypeSpeaker**|
|**ppShowTypeWindow**|

## Example

This example runs a slide show of the active presentation in a window, starting with slide two and ending with slide four. The new slide show window is placed in the upper-left corner of the screen, and its width and height are both 300 points.


```vb
With ActivePresentation.SlideShowSettings

    .RangeType = ppShowSlideRange

    .StartingSlide = 2

    .EndingSlide = 4

    .ShowType = ppShowTypeWindow

    With .Run

        .Left = 0

        .Top = 0

        .Width = 300

        .Height = 300

    End With

End With
```


## See also


#### Concepts


[SlideShowSettings Object](slideshowsettings-object-powerpoint.md)

