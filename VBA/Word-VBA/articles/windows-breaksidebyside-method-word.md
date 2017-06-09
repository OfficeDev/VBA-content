---
title: Windows.BreakSideBySide Method (Word)
keywords: vbawd10.chm157351949
f1_keywords:
- vbawd10.chm157351949
ms.prod: word
api_name:
- Word.Windows.BreakSideBySide
ms.assetid: 86e02a0d-4449-30e9-69a1-984e815711d4
ms.date: 06/08/2017
---


# Windows.BreakSideBySide Method (Word)

Ends side by side mode if two windows are in side by side mode. Returns a  **Boolean** that represents whether the method was successful.


## Syntax

 _expression_ . **BreakSideBySide**

 _expression_ Required. A variable that represents a **[Windows](windows-object-word.md)** collection.


### Return Value

Boolean


## Example

The following example ends side by side mode.


```vb
ActiveDocument.Windows.BreakSideBySide
```


## See also


#### Concepts


[Windows Collection Object](windows-object-word.md)

