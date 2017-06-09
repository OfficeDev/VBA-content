---
title: Bookmarks.ShowHidden Property (Word)
keywords: vbawd10.chm157745156
f1_keywords:
- vbawd10.chm157745156
ms.prod: word
api_name:
- Word.Bookmarks.ShowHidden
ms.assetid: 35f9a36c-ea29-93f0-1b39-c52dd3718ee8
ms.date: 06/08/2017
---


# Bookmarks.ShowHidden Property (Word)

 **True** if hidden bookmarks are included in the **Bookmarks** collection. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowHidden**

 _expression_ An expression that returns a **[Bookmarks](bookmarks-object-word.md)** collection object.


## Remarks

The  **ShowHidden** property also controls whether hidden bookmarks are listed in the **Bookmark** dialog box ( **Insert** menu). Hidden bookmarks are automatically inserted when cross-references are inserted into the document.


## Example

This example displays the  **Bookmark** dialog box with both visible and hidden bookmarks listed.


```vb
ActiveDocument.Bookmarks.ShowHidden = True 
Dialogs(wdDialogInsertBookmark).Show
```

This example displays the name of each hidden bookmark in the document. Hidden bookmarks in a Word document begin with an underscore ( _ ).




```vb
ActiveDocument.Bookmarks.ShowHidden = True 
For Each aBookmark In ActiveDocument.Bookmarks 
 If Left(aBookmark.Name, 1) = "_" Then MsgBox aBookmark.Name 
Next aBookmark
```


## See also


#### Concepts


[Bookmarks Collection Object](bookmarks-object-word.md)

