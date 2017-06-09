---
title: SlideShowSettings.RangeType Property (PowerPoint)
keywords: vbapp10.chm514014
f1_keywords:
- vbapp10.chm514014
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowSettings.RangeType
ms.assetid: 63e266b6-4898-abb1-23fe-20039a6aea78
ms.date: 06/08/2017
---


# SlideShowSettings.RangeType Property (PowerPoint)

Returns or sets the type of slide show to run. Read/write.


## Syntax

 _expression_. **RangeType**

 _expression_ A variable that represents a **SlideShowSettings** object.


## Remarks

The value of the  **RangeType** property can be one of these **PpSlideShowRangeType** constants.


||
|:-----|
|**ppShowAll**|
|**ppShowNamedSlideShow**|
|**ppShowSlideRange**|

## Example

This example runs the named slide show "Quick Show."


```vb
With ActivePresentation.SlideShowSettings

    .RangeType = ppShowNamedSlideShow

    .SlideShowName = "Quick Show"

    .Run

End With
```


## See also


#### Concepts


[SlideShowSettings Object](slideshowsettings-object-powerpoint.md)

