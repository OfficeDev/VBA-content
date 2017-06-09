---
title: Application.GoBack Method (Word)
keywords: vbawd10.chm158335304
f1_keywords:
- vbawd10.chm158335304
ms.prod: word
api_name:
- Word.Application.GoBack
ms.assetid: d1113bc7-4ad3-f4da-0442-c11f5e22b2a8
ms.date: 06/08/2017
---


# Application.GoBack Method (Word)

Moves the insertion point among the last three locations where editing occurred in the active document (the same as pressing SHIFT+F5).


## Syntax

 _expression_ . **GoBack**

 _expression_ Required. A variable that represents an **[Application](application-object-word.md)** object.


## Example

This example opens the most recently used file and then moves the insertion point to the location where editing last occurred.


```
RecentFiles(1).Open 
Application.GoBack
```


## See also


#### Concepts


[Application Object](application-object-word.md)

