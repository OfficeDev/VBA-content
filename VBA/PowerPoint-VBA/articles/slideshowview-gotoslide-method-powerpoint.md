---
title: SlideShowView.GotoSlide Method (PowerPoint)
keywords: vbapp10.chm513021
f1_keywords:
- vbapp10.chm513021
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowView.GotoSlide
ms.assetid: f733f46d-a632-02cb-3dbf-f29122fe347a
ms.date: 06/08/2017
---


# SlideShowView.GotoSlide Method (PowerPoint)

Switches to the specified slide during a slide show. You can specify whether you want the animation effects to be rerun.


## Syntax

 _expression_. **GotoSlide**( **_Index_**, **_ResetSlide_** )

 _expression_ A variable that represents a **SlideShowView** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Integer**|The number of the slide to switch to.|
| _ResetSlide_|Optional|**MsoTriState**|Whether animation effects should be rerun when returning to the first slide. See Remarks for more information.|

## Remarks

The value of the ResetSlide parameter can be one of these  **MsoTriState** constants. The default is **msoTrue**.


||
|:-----|
|**msoFalse**|
|**msoTrue**|
If you switch from one slide to another during a slide show with ResetSlide set to  **msoFalse**, when you return to the first slide, its animation picks up where it left off. If you switch from one slide to another with ResetSlide set to **msoTrue**, when you return to the first slide, its entire animation starts over.


## Example

This example switches from the current slide to the slide three in slide show window one. If you switch back to the current slide during the slide show, its entire animation will start over.


```vb
With SlideShowWindows(1).View

    .GotoSlide 3

End With
```

This example switches from the current slide to the slide three in slide show window one. If you switch back to the current slide during the slide show, its animation will pick up where it left off.




```vb
With SlideShowWindows(1).View

    .GotoSlide 3, msoFalse

End With
```


## See also


#### Concepts


[SlideShowView Object](slideshowview-object-powerpoint.md)

