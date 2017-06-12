---
title: Range.PreviousBookmarkID Property (Word)
keywords: vbawd10.chm157155637
f1_keywords:
- vbawd10.chm157155637
ms.prod: word
api_name:
- Word.Range.PreviousBookmarkID
ms.assetid: 19aab6c4-bc86-3f65-4fbc-206fdf3dbb3a
ms.date: 06/08/2017
---


# Range.PreviousBookmarkID Property (Word)

Returns the number of the last bookmark that starts before or at the same place as the specified range. Read-only  **Long** .


## Syntax

 _expression_ . **PreviousBookmarkID**

 _expression_ A variable that represents a **[Range](range-object-word.md)** object.


## Remarks

This property returns 0 (zero) if there is no corresponding bookmark


## Example

This example displays the name of the bookmark that precedes the second paragraph.


```vb
num = ActiveDocument.Paragraphs(2).Range.PreviousBookmarkID 
If num <> 0 Then MsgBox ActiveDocument.Content.Bookmarks(num).Name
```


## See also


#### Concepts


[Range Object](range-object-word.md)

