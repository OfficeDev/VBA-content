---
title: SlideShowSettings.StartingSlide Property (PowerPoint)
keywords: vbapp10.chm514005
f1_keywords:
- vbapp10.chm514005
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowSettings.StartingSlide
ms.assetid: e7afc69c-0224-b22a-fc23-bb985e710c1a
ms.date: 06/08/2017
---


# SlideShowSettings.StartingSlide Property (PowerPoint)

Returns or sets the first slide to be displayed in the specified slide show. Read/write.


## Syntax

 _expression_. **StartingSlide**

 _expression_ A variable that represents a **SlideShowSettings** object.


### Return Value

Long


## Example

This example runs a slide show of the active presentation, starting with slide two and ending with slide four.


```vb
With ActivePresentation.SlideShowSettings

    .RangeType = ppShowSlideRange

    .StartingSlide = 2

    .EndingSlide = 4

    .Run

End With
```


## See also


#### Concepts


[SlideShowSettings Object](slideshowsettings-object-powerpoint.md)

