---
title: Application.FocusInMailHeader Property (Word)
keywords: vbawd10.chm158335362
f1_keywords:
- vbawd10.chm158335362
ms.prod: word
api_name:
- Word.Application.FocusInMailHeader
ms.assetid: fba9d08b-1950-b825-5f1a-14d671181b22
ms.date: 06/08/2017
---


# Application.FocusInMailHeader Property (Word)

 **True** if the insertion point is in an e-mail header field (the To: field, for example). Read-only **Boolean** .


## Syntax

 _expression_ . **FocusInMailHeader**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


## Example

This example displays a message in the status bar if the insertion point is in an e-mail header field.


```vb
If Application.FocusInMailHeader = True Then 
 StatusBar = "Selection is in message header" 
End If
```


## See also


#### Concepts


[Application Object](application-object-word.md)

