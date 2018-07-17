---
title: Selection.Start Property (Word)
keywords: vbawd10.chm158662659
f1_keywords:
- vbawd10.chm158662659
ms.prod: word
api_name:
- Word.Selection.Start
ms.assetid: e1928372-2473-e377-4ba1-894b104fcf43
ms.date: 06/08/2017
---


# Selection.Start Property (Word)

Returns or sets the starting character position of a selection. Read/write  **Long** .


## Syntax

 _expression_ . **Start**

 _expression_ A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

 **Selection** objects have starting and ending character positions. The starting position refers to the character position closest to the beginning of the story. If this property is set to a value larger than that of the **End** property, the **End** property is set to the same value as that of **Start** property.

This property returns the starting character position relative to the beginning of the story. The main text story ( **wdMainTextStory** ) begins with character position 0 (zero). You can change the size of a selection, range, or bookmark by setting this property.


## Example

This example determines the length of the selection by comparing the starting and ending character positions.


```
SelLength = Selection.End - Selection.Start
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

