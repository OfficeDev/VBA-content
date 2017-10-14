---
title: Range.BookmarkID Property (Word)
keywords: vbawd10.chm157155636
f1_keywords:
- vbawd10.chm157155636
ms.prod: word
api_name:
- Word.Range.BookmarkID
ms.assetid: 11157160-6cd5-38d7-dc92-be14399509f4
ms.date: 06/08/2017
---


# Range.BookmarkID Property (Word)

Returns the number of the bookmark that encloses the beginning of the specified range; returns 0 (zero) if there is no corresponding bookmark. Read-only  **Long** .


## Syntax

 _expression_ . **BookmarkID**

 _expression_ A variable that represents a **[Range](range-object-word.md)** object.


## Remarks

The number or the bookmark corresponds to the position of the bookmark in the document: 1 for the first bookmark, 2 for the second one, and so on. 


## Example

This example adds a bookmark named "temp" at the beginning of the document if there is not already a bookmark set for that location.


```vb
Set myRange = ActiveDocument.Content 
myRange.Collapse Direction:=wdCollapseStart 
If myRange.BookmarkID = 0 Then 
 ActiveDocument.Bookmarks.Add Name:="temp", Range:=myRange 
End If
```


## See also


#### Concepts


[Range Object](range-object-word.md)

