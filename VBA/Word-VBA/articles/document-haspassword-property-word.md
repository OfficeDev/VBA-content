---
title: Document.HasPassword Property (Word)
keywords: vbawd10.chm158007383
f1_keywords:
- vbawd10.chm158007383
ms.prod: word
api_name:
- Word.Document.HasPassword
ms.assetid: 4234b91c-b82c-605a-5d6c-ff18aadc3689
ms.date: 06/08/2017
---


# Document.HasPassword Property (Word)

 **True** if a password is required to open the specified document. Read-only **Boolean** .


## Syntax

 _expression_ . **HasPassword**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example sets the password "kittycat" for the active document and then displays a confirmation message.


```vb
ActiveDocument.Password = "kittycat" 
If ActiveDocument.HasPassword = True Then _ 
 MsgBox "The password is set."
```


## See also


#### Concepts


[Document Object](document-object-word.md)

