---
title: PageSetup.LeftMargin Property (Word)
keywords: vbawd10.chm158400614
f1_keywords:
- vbawd10.chm158400614
ms.prod: word
api_name:
- Word.PageSetup.LeftMargin
ms.assetid: 873d6cf2-da9f-5d88-314f-9820284a54ee
ms.date: 06/08/2017
---


# PageSetup.LeftMargin Property (Word)

Returns or sets the distance (in points) between the left edge of the page and the left boundary of the body text. Read/write  **Single** .


## Syntax

 _expression_ . **LeftMargin**

 _expression_ An expression that returns a **[PageSetup](pagesetup-object-word.md)** object.


## Remarks

If the  **[MirrorMargins](pagesetup-mirrormargins-property-word.md)** property is set to **True** , the LeftMargin property controls the setting for inside margins and the **[RightMargin](pagesetup-rightmargin-property-word.md)** property controls the setting for outside margins.


## Example

This example sets the left margin to 1 inch (72 points) for the second section in the active document.


```vb
ActiveDocument.Sections(2).PageSetup.LeftMargin = 72
```


## See also


#### Concepts


[PageSetup Object](pagesetup-object-word.md)

