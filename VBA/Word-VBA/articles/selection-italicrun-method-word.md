---
title: Selection.ItalicRun Method (Word)
keywords: vbawd10.chm158663259
f1_keywords:
- vbawd10.chm158663259
ms.prod: word
api_name:
- Word.Selection.ItalicRun
ms.assetid: 0d36eff1-7308-7695-7058-be79455836ee
ms.date: 06/08/2017
---


# Selection.ItalicRun Method (Word)

Adds the italic character format to or removes it from the current run.


## Syntax

 _expression_ . **ItalicRun**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

If the run contains a mix of italic and non-italic text, this method adds the italic character format to the entire run.


## Example

This example toggles the italic formatting for the current selection.


```
Selection.ItalicRun
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

