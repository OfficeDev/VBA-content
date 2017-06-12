---
title: WebBrowserControl.ScrollLeft Property (Access)
keywords: vbaac10.chm14366
f1_keywords:
- vbaac10.chm14366
ms.prod: access
api_name:
- Access.WebBrowserControl.ScrollLeft
ms.assetid: 1526e744-8276-55bd-bd2a-b7c36cd7c3af
ms.date: 06/08/2017
---


# WebBrowserControl.ScrollLeft Property (Access)

Gets or sets the distance, in pixels, between the left edge of the  **WebBrowser** object and the leftmost portion of the content currently visible in the control. Read/write **Long**.


## Syntax

 _expression_. **ScrollLeft**

 _expression_ A variable that represents a **WebBrowserControl** object.


## Remarks

This property value equals the current horizontal offset of the content within the scrollable range. Although you can set this property to any value, if you assign a value less than 0, the property is set to 0. If you assign a value greater than the maximum value, the property is set to the maximum value.

This property is always 0 for objects that do not have scroll bars. For these objects, setting the property has no effect.


## See also


#### Concepts


[WebBrowserControl Object](webbrowsercontrol-object-access.md)

