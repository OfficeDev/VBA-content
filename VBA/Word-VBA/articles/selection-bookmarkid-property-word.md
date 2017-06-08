---
title: Selection.BookmarkID Property (Word)
keywords: vbawd10.chm158662964
f1_keywords:
- vbawd10.chm158662964
ms.prod: word
api_name:
- Word.Selection.BookmarkID
ms.assetid: f48d317c-b5ed-ff0e-4a22-13b68aa10be1
ms.date: 06/08/2017
---


# Selection.BookmarkID Property (Word)

Returns the number of the bookmark that encloses the beginning of the specified selection. Read-only  **Long** .


## Syntax

 _expression_ . **BookmarkID**

 _expression_ A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

Returns 0 (zero) if there is no corresponding bookmark. The number corresponds to the position of the bookmark in the document—1 for the first bookmark, 2 for the second one, and so on. 


## Example

This example displays the number of the bookmark that encloses the beginning of the selection.


```vb
MsgBox "Bookmark " &; Selection.BookmarkID
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

