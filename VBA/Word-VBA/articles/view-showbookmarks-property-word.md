---
title: View.ShowBookmarks Property (Word)
keywords: vbawd10.chm161808406
f1_keywords:
- vbawd10.chm161808406
ms.prod: word
api_name:
- Word.View.ShowBookmarks
ms.assetid: 20261163-6714-8361-b76d-34570868954b
ms.date: 06/08/2017
---


# View.ShowBookmarks Property (Word)

 **True** if square brackets are displayed at the beginning and end of each bookmark. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowBookmarks**

 _expression_ An expression that returns a **[View](view-object-word.md)** object.


## Example

This example displays square brackets around bookmarks in all windows.


```vb
For Each aWindow In Windows 
 aWindow.View.ShowBookmarks = True 
Next aWindow
```

This example marks the selection with a bookmark, displays square brackets around each bookmark in the active document, and then collapses the selection.




```vb
ActiveDocument.Bookmarks.Add Range:=Selection.Range, Name:="temp" 
ActiveDocument.ActiveWindow.View.ShowBookmarks = True 
Selection.Collapse Direction:=wdCollapseStart
```


## See also


#### Concepts


[View Object](view-object-word.md)

