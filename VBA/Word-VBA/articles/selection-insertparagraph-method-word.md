---
title: Selection.InsertParagraph Method (Word)
keywords: vbawd10.chm158662816
f1_keywords:
- vbawd10.chm158662816
ms.prod: word
api_name:
- Word.Selection.InsertParagraph
ms.assetid: bceda293-7294-8769-75fe-4792199439c1
ms.date: 06/08/2017
---


# Selection.InsertParagraph Method (Word)

Replaces the specified selection with a new paragraph.


## Syntax

 _expression_ . **InsertParagraph**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

After using method, the selection contains the new paragraph. If you don't want to replace the current selection, use the  **Collapse** method before using this method. You can also use the **[InsertParagraphBefore](selection-insertparagraphbefore-method-word.md)** or **[InsertParagraphAfter](selection-insertparagraphafter-method-word.md)** method to insert a new paragraph before or after a selection.


## Example

This example collapses the selection and then inserts a paragraph mark at the insertion point.


```vb
With Selection 
 .Collapse Direction:=wdCollapseStart 
 .InsertParagraph 
 .Collapse Direction:=wdCollapseEnd 
End With
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

