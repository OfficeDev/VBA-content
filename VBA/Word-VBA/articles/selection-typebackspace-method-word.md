---
title: Selection.TypeBackspace Method (Word)
keywords: vbawd10.chm158663169
f1_keywords:
- vbawd10.chm158663169
ms.prod: word
api_name:
- Word.Selection.TypeBackspace
ms.assetid: 479f2e0e-06d6-cd62-dc3e-09a5fafafbfa
ms.date: 06/08/2017
---


# Selection.TypeBackspace Method (Word)

Deletes the character preceding a collapsed selection (an insertion point).


## Syntax

 _expression_ . **TypeBackspace**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

This method corresponds to functionality of the BACKSPACE key. If the selection isn't collapsed to an insertion point, the selection is deleted.


## Example

This example deletes the character preceding the insertion point (the collapsed selection).


```vb
With Selection 
 .Collapse Direction:=wdCollapseEnd 
 .TypeBackspace 
End With
```

This example extends the selection to the end of the current paragraph (including the paragraph mark) and then deletes the selection.




```vb
With Selection 
 .EndOf Unit:=wdParagraph, Extend:=wdExtend 
 .TypeBackspace 
End With
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

