---
title: ShapeRange.LockAnchor Property (Word)
keywords: vbawd10.chm162857262
f1_keywords:
- vbawd10.chm162857262
ms.prod: word
api_name:
- Word.ShapeRange.LockAnchor
ms.assetid: 63137738-47cb-bb2a-eb3a-25c421de298a
ms.date: 06/08/2017
---


# ShapeRange.LockAnchor Property (Word)

 **True** if the anchor for the specified **ShapeRange** object is locked to the anchoring range. Read/write **Long** .


## Syntax

 _expression_ . **LockAnchor**

 _expression_ Required. A variable that represents a **[ShapeRange](shaperange-object-word.md)** object.


## Remarks

When a range of shapes has a locked anchor, you cannot move the shapes' anchor by dragging it. The anchor does not move as the shape is moved.

A  **ShapeRange** object is anchored to a range of text, but you can position it anywhere on the page. If the range of shapes is anchored to the beginning of the first paragraph that contains the anchoring range, the shapes always remain on the same page as the anchor.


## See also


#### Concepts


[ShapeRange Collection Object](shaperange-object-word.md)

