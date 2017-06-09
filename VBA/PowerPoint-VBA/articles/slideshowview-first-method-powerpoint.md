---
title: SlideShowView.First Method (PowerPoint)
keywords: vbapp10.chm513017
f1_keywords:
- vbapp10.chm513017
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowView.First
ms.assetid: 5f360832-2deb-b3df-7b55-5a3c964d0057
ms.date: 06/08/2017
---


# SlideShowView.First Method (PowerPoint)

Sets the specified slide show view to display the first slide in the presentation.


## Syntax

 _expression_. **First**

 _expression_ A variable that represents a **SlideShowView** object.


### Return Value

Nothing


## Remarks

If you use the  **First** method to switch from one slide to another during a slide show, when you return to the original slide, its animation picks up where it left off.


## Example

This example sets slide show window one to display the first slide in the presentation.


```
SlideShowWindows(1).View.First
```


## See also


#### Concepts


[SlideShowView Object](slideshowview-object-powerpoint.md)

