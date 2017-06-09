---
title: Bookmark.End Property (Word)
keywords: vbawd10.chm157810692
f1_keywords:
- vbawd10.chm157810692
ms.prod: word
api_name:
- Word.Bookmark.End
ms.assetid: 05531b0d-b05e-0010-9ff8-ba6d90de560d
ms.date: 06/08/2017
---


# Bookmark.End Property (Word)

Returns or sets the ending character position of a selection, range, or bookmark. Read/write  **Long** .


## Syntax

 _expression_ . **End**

 _expression_ A variable that represents a **[Bookmark](bookmark-object-word.md)** object.


## Remarks

If this property is set to a value smaller than the  **[Start](bookmark-start-property-word.md)** property, the **Start** property is set to the same value (that is, the **Start** and **End** properties are equal).

This property returns the ending character position relative to the beginning of the story. The main document story (wdMainTextStory) begins with character position 0 (zero). You can change the size of a bookmark by setting this property.


## Example

This example compares the ending position of the "temp" bookmark with the starting position of the "begin" bookmark.


```vb
Set Book1 = ActiveDocument.Bookmarks("begin") 
Set Book2 = ActiveDocument.Bookmarks("temp") 
If Book2.End > Book1.Start Then Book1.Select
```


## See also


#### Concepts


[Bookmark Object](bookmark-object-word.md)

