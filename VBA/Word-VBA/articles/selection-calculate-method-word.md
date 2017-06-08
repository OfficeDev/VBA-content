---
title: Selection.Calculate Method (Word)
keywords: vbawd10.chm158662828
f1_keywords:
- vbawd10.chm158662828
ms.prod: word
api_name:
- Word.Selection.Calculate
ms.assetid: a4e7ef08-8442-0579-e738-e4f53ee62d62
ms.date: 06/08/2017
---


# Selection.Calculate Method (Word)

Calculates a mathematical expression within a selection. Returns the result as a  **Single** .


## Syntax

 _expression_ . **Calculate**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Example

This example calculates the selected mathematical expression and displays the result.


```vb
MsgBox "And the answer is... " &; Selection.Calculate
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

