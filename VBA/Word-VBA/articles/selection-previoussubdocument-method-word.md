---
title: Selection.PreviousSubdocument Method (Word)
keywords: vbawd10.chm158663171
f1_keywords:
- vbawd10.chm158663171
ms.prod: word
api_name:
- Word.Selection.PreviousSubdocument
ms.assetid: 135932fa-c165-56d1-97c7-f04fd7108ab2
ms.date: 06/08/2017
---


# Selection.PreviousSubdocument Method (Word)

Moves the selection to the previous subdocument.


## Syntax

 _expression_ . **PreviousSubdocument**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

If there isn't another subdocument, an error occurs.


## Example

This example switches the active document to master document view and selects the previous subdocument.


```vb
If ActiveDocument.Subdocuments.Count >= 1 Then 
 ActiveDocument.ActiveWindow.View.Type = wdMasterView 
 Selection.EndKey Unit:=wdStory, Extend:=wdMove 
 Selection.PreviousSubdocument 
End If
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

