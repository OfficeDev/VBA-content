---
title: Options.DefaultEPostageApp Property (Word)
keywords: vbawd10.chm162988474
f1_keywords:
- vbawd10.chm162988474
ms.prod: word
api_name:
- Word.Options.DefaultEPostageApp
ms.assetid: 1d039201-2e86-7f8b-9732-da1d13a12cf0
ms.date: 06/08/2017
---


# Options.DefaultEPostageApp Property (Word)

Sets or returns a  **String** that represents the path and file name of the default electronic postage application. Read/write.


## Syntax

 _expression_ . **DefaultEPostageApp**

 _expression_ A variable that represents a **[Options](options-object-word.md)** object.


## Example

This example specifies the path and file name for the default electronic postage application.


```vb
Sub DefaultEPostage() 
 Application.Options.DefaultEPostageApp = "C:\MyApp\EPostage.exe" 
End Sub
```


## See also


#### Concepts


[Options Object](options-object-word.md)

