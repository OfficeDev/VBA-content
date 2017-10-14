---
title: ScrollBar.ProportionalThumb Property (Outlook Forms Script)
keywords: olfm10.chm2001750
f1_keywords:
- olfm10.chm2001750
ms.prod: outlook
ms.assetid: 3238c848-3279-9a3b-a576-136d9f1ddf28
ms.date: 06/08/2017
---


# ScrollBar.ProportionalThumb Property (Outlook Forms Script)

Returns or sets a  **Boolean** that specifies whether the size of the scroll box is proportional to the scrolling region or fixed. Read/write.


## Syntax

 _expression_. **ProportionalThumb**

 _expression_A variable that represents a  **ScrollBar** object.


## Remarks

 **True** if the scroll box is proportional in size to the scrolling region (default). **False** if the scroll box is a fixed size.

The size of a proportional scroll box graphically represents the percentage of the object that is visible in the window. For example, if 75 percent of an object is visible, the scroll box covers three-fourths of the scrolling region in the scroll bar.

If the scroll box is a fixed size, the system determines its size based on the height and width of the scroll bar.


