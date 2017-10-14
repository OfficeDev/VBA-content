---
title: SlideShowView.PointerColor Property (PowerPoint)
keywords: vbapp10.chm513012
f1_keywords:
- vbapp10.chm513012
ms.prod: powerpoint
api_name:
- PowerPoint.SlideShowView.PointerColor
ms.assetid: 29f4c5e0-0927-1dbb-7bc9-b147ae38ff88
ms.date: 06/08/2017
---


# SlideShowView.PointerColor Property (PowerPoint)

Returns a  **ColorFormat** object that represents the pointer color for the specified presentation during one slide show. Read-only.


## Syntax

 _expression_. **PointerColor**

 _expression_ A variable that represents a **SlideShowView** object.


### Return Value

ColorFormat


## Remarks

As soon as the slide show is finished, the color reverts to the default color for the presentation. 

To change the pointer to a pen, set the  **[PointerType](slideshowview-pointertype-property-powerpoint.md)** property to **ppSlideShowPointerPen**.


## See also


#### Concepts


[SlideShowView Object](slideshowview-object-powerpoint.md)

