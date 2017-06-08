---
title: FormField.Cut Method (Word)
keywords: vbawd10.chm153616486
f1_keywords:
- vbawd10.chm153616486
ms.prod: word
api_name:
- Word.FormField.Cut
ms.assetid: 92b8862d-6463-0bbd-cffd-8e76f5add5b4
ms.date: 06/08/2017
---


# FormField.Cut Method (Word)

Removes the specified form field from the document and places it on the Clipboard.


## Syntax

 _expression_ . **Cut**

 _expression_ Required. A variable that represents a **[FormField](formfield-object-word.md)** object.


## Example

This example deletes the first field in the active document and pastes the field at the insertion point.


```vb
If ActiveDocument.Fields.Count >= 1 Then 
 ActiveDocument.Fields(1).Cut 
 Selection.Collapse Direction:=wdCollapseEnd 
 Selection.Paste 
End If
```

This example deletes the first word in the first paragraph and pastes the word at the end of the paragraph.




```vb
With ActiveDocument.Paragraphs(1).Range 
 .Words(1).Cut 
 .Collapse Direction:=wdCollapseEnd 
 .Move Unit:=wdCharacter, Count:=-1 
 .Paste 
End With
```

This example deletes the contents of the selection and pastes them into a new document.




```vb
If Selection.Type = wdSelectionNormal Then 
 Selection.Cut 
 Documents.Add.Content.Paste 
End If
```


## See also


#### Concepts


[FormField Object](formfield-object-word.md)

