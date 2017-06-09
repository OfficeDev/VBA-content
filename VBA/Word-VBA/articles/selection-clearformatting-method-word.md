---
title: Selection.ClearFormatting Method (Word)
keywords: vbawd10.chm158663665
f1_keywords:
- vbawd10.chm158663665
ms.prod: word
api_name:
- Word.Selection.ClearFormatting
ms.assetid: 66c2f088-5d35-f8b0-10e5-2faa0db14d7f
ms.date: 06/08/2017
---


# Selection.ClearFormatting Method (Word)

Removes text and paragraph formatting from a selection.


## Syntax

 _expression_ . **ClearFormatting**

 _expression_ A variable that represents a **[Selection](selection-object-word.md)** object.


## Example

This example removes all text and paragraph formatting from the active document.


```vb
Sub ClrFmtg() 
 ActiveDocument.Select 
 Selection.ClearFormatting 
End Sub
```

This example removes all text and paragraph formatting from the second through the fourth paragraphs of the active document.




```vb
Sub ClrFmtg2() 
 ActiveDocument.Range(Start:=ActiveDocument.Paragraphs(2).Range.Start, _ 
 End:=ActiveDocument.Paragraphs(4).Range.End).Select 
 Selection.ClearFormatting 
End Sub
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

