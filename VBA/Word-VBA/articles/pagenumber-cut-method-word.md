---
title: PageNumber.Cut Method (Word)
keywords: vbawd10.chm159842406
f1_keywords:
- vbawd10.chm159842406
ms.prod: word
api_name:
- Word.PageNumber.Cut
ms.assetid: 20813c72-2a09-8115-dbfe-ed738dbdbe7c
ms.date: 06/08/2017
---


# PageNumber.Cut Method (Word)

Removes the specified object from the document and places it on the Clipboard.


## Syntax

 _expression_ . **Cut**

 _expression_ Required. A variable that represents a **[PageNumber](pagenumber-object-word.md)** object.


## Remarks

If expression returns a  **Range** or **Selection** object, the contents of the object are moved to the Clipboard but the collapsed object remains in the document.


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


[PageNumber Object](pagenumber-object-word.md)

