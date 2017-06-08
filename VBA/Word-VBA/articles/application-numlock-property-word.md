---
title: Application.NumLock Property (Word)
keywords: vbawd10.chm158335025
f1_keywords:
- vbawd10.chm158335025
ms.prod: word
api_name:
- Word.Application.NumLock
ms.assetid: 0c20c000-2df9-1483-91be-cacf1abe0ff0
ms.date: 06/08/2017
---


# Application.NumLock Property (Word)

Returns the state of the NUM LOCK key.  **True** if the keys on the numeric keypad insert numbers, **False** if the keys move the insertion point. Read-only **Boolean** .


## Syntax

 _expression_ . **NumLock**

 _expression_ An expression that returns an **[Application](application-object-word.md)** object.


## Example

This example returns the current state of the NUM LOCK key.


```
theState = Application.NumLock
```


## See also


#### Concepts


[Application Object](application-object-word.md)

