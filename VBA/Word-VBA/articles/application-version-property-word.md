---
title: Application.Version Property (Word)
keywords: vbawd10.chm158335000
f1_keywords:
- vbawd10.chm158335000
ms.prod: word
api_name:
- Word.Application.Version
ms.assetid: 7bdd0acc-1ed0-677c-f973-99a9199e030b
ms.date: 06/08/2017
---


# Application.Version Property (Word)

Returns the Microsoft Word version number. Read-only  **String** .


## Syntax

 _expression_ . **Version**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


## Example

This example displays the Word version number in a message box.


```
Msgbox "The version of Word is " &; Application.Version
```


## See also


#### Concepts


[Application Object](application-object-word.md)

