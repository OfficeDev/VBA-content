---
title: Shape.Anchor Property (Word)
keywords: vbawd10.chm161481205
f1_keywords:
- vbawd10.chm161481205
ms.prod: word
api_name:
- Word.Shape.Anchor
ms.assetid: a2889d2a-4b47-cf27-a9ef-b9fe479b7929
ms.date: 06/08/2017
---


# Shape.Anchor Property (Word)

Returns a  **Range** object that represents the anchoring range for the specified shape or shape range. Read-only.


## Syntax

 _expression_ . **Anchor**

 _expression_ A variable that represents a **[Shape](shape-object-word.md)** object.


## Remarks

All  **Shape** objects are anchored to a range of text but can be positioned anywhere on the page that contains the anchor. If you specify the anchoring range when you create a shape, the anchor is positioned at the beginning of the first paragraph that contains the anchoring range. If you don't specify the anchoring range, the anchoring range is selected automatically and the shape is positioned relative to the top and left edges of the page.

The shape will always remain on the same page as its anchor. If the  **LockAnchor** property for the shape is set to **True** , you cannot drag the anchor from its position on the page.


## Example

This example selects the paragraph that the first shape in the active document is anchored to.


```vb
ActiveDocument.Shapes(1).Anchor.Paragraphs(1).Range.Select
```


## See also


#### Concepts


[Shape Object](shape-object-word.md)

