---
title: Selection.PasteFormat Method (Word)
keywords: vbawd10.chm158663166
f1_keywords:
- vbawd10.chm158663166
ms.prod: word
api_name:
- Word.Selection.PasteFormat
ms.assetid: 5c8a69fa-4d07-619c-950a-5ff11fa99003
ms.date: 06/08/2017
---


# Selection.PasteFormat Method (Word)

Applies formatting copied with the  **CopyFormat** method to the selection.


## Syntax

 _expression_ . **PasteFormat**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

If a paragraph mark was selected when the  **CopyFormat** method was used, Word applies paragraph formatting in addition to character formatting.


## Example

This example copies the paragraph and character formatting from the first paragraph in the selection to the next paragraph in the selection.


```vb
With Selection 
 .Paragraphs(1).Range.Select 
 .CopyFormat 
 .Paragraphs(1).Next.Range.Select 
 .PasteFormat 
End With
```

This example collapses the selection and copies the character formatting to the next word.




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

