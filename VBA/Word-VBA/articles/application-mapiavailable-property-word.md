---
title: Application.MAPIAvailable Property (Word)
keywords: vbawd10.chm158335074
f1_keywords:
- vbawd10.chm158335074
ms.prod: word
api_name:
- Word.Application.MAPIAvailable
ms.assetid: 2cb2fc8c-1ef6-98b8-fa72-0705637ad3ac
ms.date: 06/08/2017
---


# Application.MAPIAvailable Property (Word)

 **True** if MAPI is installed. Read-only **Boolean** .


## Syntax

 _expression_ . **MAPIAvailable**

 _expression_ An expression that returns an **[Application](application-object-word.md)** object.


## Example

This example displays a message if MAPI is installed.


```vb
If Application.MAPIAvailable = True Then 
 MsgBox "MAPI is available" 
End If
```


## See also


#### Concepts


[Application Object](application-object-word.md)

