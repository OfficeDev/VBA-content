---
title: Document.WritePassword Property (Word)
keywords: vbawd10.chm158007382
f1_keywords:
- vbawd10.chm158007382
ms.prod: word
api_name:
- Word.Document.WritePassword
ms.assetid: e3353e68-1196-d896-d978-2c49ceca2940
ms.date: 06/08/2017
---


# Document.WritePassword Property (Word)

Sets a password for saving changes to the specified document. Write-only  **String** .


## Syntax

 _expression_ . **WritePassword**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks


 **Important**  Avoid using hard-coded passwords in your applications. If a password is required in a procedure, request the password from the user, store it in a variable, and then use the variable in your code. For recommended best practices on how to do this, see [Security Notes for Microsoft Office Solution Developers](https://msdn.microsoft.com/en-us/library/office/ff860261.aspx). 


## Example

If the active document isn't already protected against saving changes, this example sets "secret" as the write password for the document.


```vb
Set myDoc = ActiveDocument 
If myDoc.WriteReserved = False Then myDoc.WritePassword = "secret"
```


## See also


#### Concepts


[Document Object](document-object-word.md)

