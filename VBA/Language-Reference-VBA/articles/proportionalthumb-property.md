---
title: ProportionalThumb Property
keywords: fm20.chm2001750
f1_keywords:
- fm20.chm2001750
ms.prod: office
api_name:
- Office.ProportionalThumb
ms.assetid: da2890ca-12b9-8d91-5e94-9c86492f0101
ms.date: 06/08/2017
---


# ProportionalThumb Property



Specifies whether the size of the scroll box is proportional to the scrolling region or fixed.
 **Syntax**
 _object_. **ProportionalThumb** [= _Boolean_ ]
The  **ProportionalThumb** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Boolean_|Optional. Whether the scroll box is proportional or fixed.|
 **Settings**
The settings for  _Boolean_ are:


|**Value**|**Description**|
|:-----|:-----|
|**True**|The scroll box is proportional in size to the scrolling region (default).|
|**False**|The scroll box is a fixed size.|
 **Remarks**
The size of a proportional scroll box graphically represents the percentage of the object that is visible in the window. For example, if 75 percent of an object is visible, the scroll box covers three-fourths of the scrolling region in the scroll bar.
If the scroll box is a fixed size, the system determines its size based on the height and width of the scroll bar.

