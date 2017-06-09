---
title: Selection.CopyFormat Method (Word)
keywords: vbawd10.chm158663165
f1_keywords:
- vbawd10.chm158663165
ms.prod: word
api_name:
- Word.Selection.CopyFormat
ms.assetid: ef892e50-2ff1-3ab0-1112-cf6d268a1103
ms.date: 06/08/2017
---


# Selection.CopyFormat Method (Word)

Copies the character formatting of the first character in the selected text.


## Syntax

 _expression_ . **CopyFormat**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

If a paragraph mark is selected, Word copies paragraph formatting in addition to character formatting. You can apply the copied formatting to another selection by using the  **[PasteFormat](selection-pasteformat-method-word.md)** method.


## Example

This example copies the formatting of the first paragraph to the second paragraph in the active document.


```vb
ActiveDocument.Paragraphs(1).Range.Select 
Selection.CopyFormat 
ActiveDocument.Paragraphs(2).Range.Select 
Selection.PasteFormat
```

This example collapses the selection and copies its character formatting to the next word.




```vb
With Selection 
 .Collapse Direction:=wdCollapseStart 
 .CopyFormat 
 .Next(Unit:=wdWord, Count:=1).Select 
 .PasteFormat 
End With
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

