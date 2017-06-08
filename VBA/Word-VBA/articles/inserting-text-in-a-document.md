---
title: Inserting Text in a Document
keywords: vbawd10.chm5211121
f1_keywords:
- vbawd10.chm5211121
ms.prod: word
ms.assetid: 4903a9aa-6923-da80-fcc0-f0e2defcb77a
ms.date: 06/08/2017
---


# Inserting Text in a Document

Use the  **InsertBefore**method or the  **InsertAfter**method of the  **[Selection](selection-object-word.md)** object or the  **[Range](range-object-word.md)** object to insert text before or after a selection or range of text. The following example inserts text at the end of the active document.


```vb
Sub InsertTextAtEndOfDocument() 
 ActiveDocument.Content.InsertAfter Text:=" The end." 
End Sub
```


The following example inserts text before the selection.




```vb
Sub AddTextBeforeSelection() 
 Selection.InsertBefore Text:="new text " 
End Sub
```

After using the  **InsertBefore** method or the **InsertAfter** method, the **Range** or **Selection** expands to include the new text. Use the **Collapse**method to collapse a  **Selection** or **Range** to the beginning or ending point.

