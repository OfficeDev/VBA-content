---
title: Application.CheckLanguage Property (Word)
keywords: vbawd10.chm158335088
f1_keywords:
- vbawd10.chm158335088
ms.prod: word
api_name:
- Word.Application.CheckLanguage
ms.assetid: 25c2a119-2cae-48e4-1d54-cafc763b90fa
ms.date: 06/08/2017
---


# Application.CheckLanguage Property (Word)

 **True** if Microsoft Word automatically detects the language you are using as you type. Read/write **Boolean** .


## Syntax

 _expression_ . **CheckLanguage**

 _expression_ An expression that returns an **[Application](application-object-word.md)** object.


## Remarks

If you haven't set up Word for multilingual editing, the  **CheckLanguage** property always returns **False** .


## Example

This example checks to see if automatic language detection has been activated.


```vb
If Application.CheckLanguage = True Then 
 MsgBox "Automatic language detection is activated!" 
End If
```


## See also


#### Concepts


[Application Object](application-object-word.md)

