---
title: ShapeRange.Anchor Property (Word)
keywords: vbawd10.chm162857264
f1_keywords:
- vbawd10.chm162857264
ms.prod: word
api_name:
- Word.ShapeRange.Anchor
ms.assetid: ee0b66e6-7385-bf61-79a3-14d874324f58
ms.date: 06/08/2017
---


# ShapeRange.Anchor Property (Word)

Returns a  **Range** object that represents the anchoring range for the specified shape range. Read-only.


## Syntax

 _expression_ . **Anchor**

 _expression_ A variable that represents a **[ShapeRange](shaperange-object-word.md)** object.


## Remarks

If you use this property on a  **ShapeRange** object that contains more than one shape, an error occurs.

All  **Shape** objects are anchored to a range of text but can be positioned anywhere on the page that contains the anchor. If you specify the anchoring range when you create a shape, the anchor is positioned at the beginning of the first paragraph that contains the anchoring range. If you don't specify the anchoring range, the anchoring range is selected automatically and the shape is positioned relative to the top and left edges of the page.

The shape will always remain on the same page as its anchor. If the  **LockAnchor** property for the shape is set to **True** , you cannot drag the anchor from its position on the page.


## See also


#### Concepts


[ShapeRange Collection Object](shaperange-object-word.md)

