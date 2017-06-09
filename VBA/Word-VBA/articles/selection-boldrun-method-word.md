---
title: Selection.BoldRun Method (Word)
keywords: vbawd10.chm158663258
f1_keywords:
- vbawd10.chm158663258
ms.prod: word
api_name:
- Word.Selection.BoldRun
ms.assetid: 0998afe2-dcd9-c1e4-9614-a1af4c6bbeaf
ms.date: 06/08/2017
---


# Selection.BoldRun Method (Word)

Adds the bold character format to or removes it from the current run.


## Syntax

 _expression_ . **BoldRun**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

 If the run contains a mix of bold and non-bold text, this method adds the bold character format to the entire run.


## Example

This example toggles the bold formatting for the current selection.


```
Selection.BoldRun
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

