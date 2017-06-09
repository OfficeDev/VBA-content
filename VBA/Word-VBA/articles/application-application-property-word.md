---
title: Application.Application Property (Word)
keywords: vbawd10.chm158335976
f1_keywords:
- vbawd10.chm158335976
ms.prod: word
api_name:
- Word.Application.Application
ms.assetid: 90d01c40-6b41-7b61-d989-6a864e6c2ca3
ms.date: 06/08/2017
---


# Application.Application Property (Word)

Returns an  **[Application](application-object-word.md)** object that represents the Microsoft Word application.


## Syntax

 _expression_ . **Application**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


## Example

This example displays scroll bars, screen tips, and the status bar for Microsoft Word.


```vb
With Application 
 .DisplayScrollBars = True 
 .DisplayScreenTips = True 
 .DisplayStatusBar = True 
End With
```


## See also


#### Concepts


[Application Object](application-object-word.md)

