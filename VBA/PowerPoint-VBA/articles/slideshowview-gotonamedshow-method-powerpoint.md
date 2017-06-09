---
title: SlideShowView.GotoNamedShow Method (PowerPoint)
keywords: vbapp10.chm513022
f1_keywords:
- vbapp10.chm513022
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowView.GotoNamedShow
ms.assetid: 7e26b77f-bb7b-fd32-eabf-bc8f568e5c62
ms.date: 06/08/2017
---


# SlideShowView.GotoNamedShow Method (PowerPoint)

Switches to the specified custom, or named, slide show during another slide show. When the slide show advances from the current slide, the next slide displayed will be the next one in the specified custom slide show, not the next one in current slide show.


## Syntax

 _expression_. **GotoNamedShow**( **_SlideShowName_** )

 _expression_ A variable that represents a **SlideShowView** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SlideShowName_|Required|**String**|The name of the custom slide show to be switched to.|

## Example

This example redefines the slide show running in slide show window one to include only the slides in the custom slide show named "Quick Show."


```
SlideShowWindows(1).View.GotoNamedShow "Quick Show"
```


## See also


#### Concepts


[SlideShowView Object](slideshowview-object-powerpoint.md)

