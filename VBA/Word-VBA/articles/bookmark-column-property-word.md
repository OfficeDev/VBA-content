---
title: Bookmark.Column Property (Word)
keywords: vbawd10.chm157810693
f1_keywords:
- vbawd10.chm157810693
ms.prod: word
api_name:
- Word.Bookmark.Column
ms.assetid: 09c819bf-e7cd-caa0-106f-8a149b4c42f8
ms.date: 06/08/2017
---


# Bookmark.Column Property (Word)

 **True** if the specified bookmark is a table column. Read-only **Boolean** .


## Syntax

 _expression_ . **Column**

 _expression_ Required. A variable that represents a **[Bookmark](bookmark-object-word.md)** object.


## Example

This example creates a table with a bookmark and then displays a message box that confirms that the bookmark is a table column.


```vb
Dim docNew As Document 
Dim tableNew As Table 
Dim rangeCell As Range 
 
Set docNew = Documents.Add 
Set tableNew = docNew.Tables.Add(Selection.Range, 3, 5) 
Set rangeCell = tableNew.Cell(3,5).Range 
 
rangeCell.InsertAfter "Cell(3,5)" 
docNew.Bookmarks.Add Name:="BKMK_Cell35", Range:=rangeCell 
MsgBox docNew.Bookmarks(1).Column
```


## See also


#### Concepts


[Bookmark Object](bookmark-object-word.md)

