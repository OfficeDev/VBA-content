---
title: Range.Paste Method (Word)
keywords: vbawd10.chm157155449
f1_keywords:
- vbawd10.chm157155449
ms.prod: word
api_name:
- Word.Range.Paste
ms.assetid: 06621016-de31-c61b-a9d0-6544b2d7e0a4
ms.date: 06/08/2017
---


# Range.Paste Method (Word)

Inserts the contents of the Clipboard at the specified range.


## Syntax

 _expression_ . **Paste**

 _expression_ Required. A variable that represents a **[Range](range-object-word.md)** object.


## Remarks

If you don't want to replace the contents of the range, use the  **Collapse** method before using this method.

When you use this method with a  **Range** object, the range expands to include the contents of the Clipboard.


## Example

This example copies and pastes the first table in the active document into a new document.


```vb
If ActiveDocument.Tables.Count >= 1 Then 
 ActiveDocument.Tables(1).Range.Copy 
 Documents.Add.Content.Paste 
End If
```

This example copies the selection and pastes it at the end of the document.




```vb
If Selection.Type <> wdSelectionIP Then 
 Selection.Copy 
 Set Range2 = ActiveDocument.Content 
 Range2.Collapse Direction:=wdCollapseEnd 
 Range2.Paste 
End If
```


## See also


#### Concepts


[Range Object](range-object-word.md)

