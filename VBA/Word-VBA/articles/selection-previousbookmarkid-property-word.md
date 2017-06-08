---
title: Selection.PreviousBookmarkID Property (Word)
keywords: vbawd10.chm158662965
f1_keywords:
- vbawd10.chm158662965
ms.prod: word
api_name:
- Word.Selection.PreviousBookmarkID
ms.assetid: 33d7490d-1b48-81a1-a7d5-9154c1d92230
ms.date: 06/08/2017
---


# Selection.PreviousBookmarkID Property (Word)

Returns the number of the last bookmark that starts before or at the same place as the specified selection or range; returns 0 (zero) if there is no corresponding bookmark. Read-only  **Long** .


## Syntax

 _expression_ . **PreviousBookmarkID**

 _expression_ A variable that represents a **[Selection](selection-object-word.md)** object.


## Example

This example selects the previous bookmark in the active document.


```vb
num = Selection.PreviousBookmarkID 
If num <> 0 Then ActiveDocument.Content.Bookmarks(num).Select
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

