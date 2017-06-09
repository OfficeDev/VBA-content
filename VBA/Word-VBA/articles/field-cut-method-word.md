---
title: Field.Cut Method (Word)
keywords: vbawd10.chm154075242
f1_keywords:
- vbawd10.chm154075242
ms.prod: word
api_name:
- Word.Field.Cut
ms.assetid: 594b6538-fd90-a969-decd-1468b9ba0c03
ms.date: 06/08/2017
---


# Field.Cut Method (Word)

Removes the specified field from the document and places it on the Clipboard.


## Syntax

 _expression_ . **Cut**

 _expression_ Required. A variable that represents a **[Field](field-object-word.md)** object.


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


[Field Object](field-object-word.md)

