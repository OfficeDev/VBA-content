---
title: Application.BrowseExtraFileTypes Property (Word)
keywords: vbawd10.chm158335084
f1_keywords:
- vbawd10.chm158335084
ms.prod: word
api_name:
- Word.Application.BrowseExtraFileTypes
ms.assetid: e411bb7a-d40f-1bda-5424-6202ba346717
ms.date: 06/08/2017
---


# Application.BrowseExtraFileTypes Property (Word)

Set this property to "text/html" to allow hyperlinked HTML files to be opened in Microsoft Word (instead of the default Internet browser). Read/write  **String** .


## Syntax

 _expression_ . **BrowseExtraFileTypes**

 _expression_ A variable that represents an **[Application](application-object-word.md)** object.


## Example

This example allows hyperlinked HTML files to be opened in Word (instead of the default Internet browser).


```vb
Application.BrowseExtraFileTypes = "text/html"
```


## See also


#### Concepts


[Application Object](application-object-word.md)

