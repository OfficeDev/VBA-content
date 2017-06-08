---
title: Selection.EscapeKey Method (Word)
keywords: vbawd10.chm158663162
f1_keywords:
- vbawd10.chm158663162
ms.prod: word
api_name:
- Word.Selection.EscapeKey
ms.assetid: a498cf00-a3dc-b084-79ae-c31d6f4e5e27
ms.date: 06/08/2017
---


# Selection.EscapeKey Method (Word)

Cancels a mode such as extend or column select (equivalent to pressing the ESC key).


## Syntax

 _expression_ . **EscapeKey**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


## Example

This example turns on and then cancels extend mode.


```vb
With Selection 
 .ExtendMode = True 
 .EscapeKey 
End With
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

