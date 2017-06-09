---
title: SlideShowView.LaserPointerEnabled Property (PowerPoint)
keywords: vbapp10.chm513038
f1_keywords:
- vbapp10.chm513038
ms.assetid: 9ba56542-a2bf-28d2-9609-50f9a4144c91
ms.date: 06/08/2017
ms.prod: powerpoint
---


# SlideShowView.LaserPointerEnabled Property (PowerPoint)

Returns  **true** if the current slide show pointer is a laser pointer. This property is applicable only while the slide show is running. Read/write.

This property allows a user to programmatically query and set the state of the pointer shown during slide show. The property will return false for all other pointer types. Users can also change the state of the current pointer by setting this property to  **true** to turn on the laser pointer or **false** to turn off the laser pointer.

## Syntax

 _expression_. **LaserPointerEnabled**

 _expression_ A variable that represents a **SlideShowView** object.


### Return Value

Boolean


## See also


#### Concepts


[SlideShowView Object](slideshowview-object-powerpoint.md)

