---
title: Selection.Bookmarks Property (Word)
keywords: vbawd10.chm158662731
f1_keywords:
- vbawd10.chm158662731
ms.prod: word
api_name:
- Word.Selection.Bookmarks
ms.assetid: 32e25786-512a-5bee-4ba6-42c801b49176
ms.date: 06/08/2017
---


# Selection.Bookmarks Property (Word)

Returns a  **[Bookmarks](bookmarks-object-word.md)** collection that represents all the bookmarks in a document, range, or selection. Read-only.


## Syntax

 _expression_ . **Bookmarks**

 _expression_ A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example applies bold formatting to the first range of bookmarked text in the selection.


```vb
If Selection.Bookmarks.Count >= 1 Then 
 Selection.Bookmarks(1).Range.Bold = True 
End If
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

