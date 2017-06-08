---
title: Editors Object (Word)
keywords: vbawd10.chm140
f1_keywords:
- vbawd10.chm140
ms.prod: word
api_name:
- Word.Editors
ms.assetid: acce718a-e3c1-deac-8b7f-fd8a5a9e47c6
ms.date: 06/08/2017
---


# Editors Object (Word)

A collection of  **Editor** objects that represents a collection of users or groups of users who have been given specific permissions to edit portions of a document.


## Remarks

Use the  **Add** method to give a specified user or group permission to modify a range or selection within a document. The following example gives the current user editing permission to modify the active selection.


```vb
Dim objEditor As Editor 
 
Set objEditor = Selection.Editors.Add(wdEditorCurrent)
```


## See also


#### Other resources



[Word Object Model Reference](http://msdn.microsoft.com/library/be452561-b436-bb9b-6f94-3faa9a74a6fd%28Office.15%29.aspx)

