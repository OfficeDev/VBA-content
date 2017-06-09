---
title: SlideShowView.EndNamedShow Method (PowerPoint)
keywords: vbapp10.chm513023
f1_keywords:
- vbapp10.chm513023
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowView.EndNamedShow
ms.assetid: 1b829558-a729-8aa1-c260-8b7410501153
ms.date: 06/08/2017
---


# SlideShowView.EndNamedShow Method (PowerPoint)

Switches from running a custom, or named, slide show to running the entire presentation of which the custom show is a subset. When the slide show advances from the current slide, the next slide displayed will be the next one in the entire presentation, not the next one in the custom slide show.


## Syntax

 _expression_. **EndNamedShow**

 _expression_ A variable that represents a **SlideShowView** object.


## Example

If a custom slide show is currently running in slide show window one, this example redefines the slide show to include all the slides in the presentation from which the slides in the custom show were selected.


```
SlideShowWindows(1).View.EndNamedShow
```


## See also


#### Concepts


[SlideShowView Object](slideshowview-object-powerpoint.md)

