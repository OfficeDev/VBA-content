---
title: Document.Password Property (Word)
keywords: vbawd10.chm158007381
f1_keywords:
- vbawd10.chm158007381
ms.prod: word
api_name:
- Word.Document.Password
ms.assetid: 243f1735-5367-4ac9-5643-624ccf501abe
ms.date: 06/08/2017
---


# Document.Password Property (Word)

Sets a password that must be supplied to open the specified document. Write-only  **String** .


## Syntax

 _expression_ . **Password**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks


 **Important**  Avoid using hard-coded passwords in your applications. If a password is required in a procedure, request the password from the user, store it in a variable, and then use the variable in your code. For recommended best practices on how to do this, see [Security Notes for Microsoft Office Solution Developers](https://msdn.microsoft.com/en-us/library/office/ff860261.aspx). 


## Example

This example opens Earnings.doc, sets a password for it, and then closes the document.


```vb
Set myDoc = Documents _ 
 .Open(FileName:="C:\My Documents\Earnings.doc") 
myDoc.Password = strPassword 
myDoc.Close
```


## See also


#### Concepts


[Document Object](document-object-word.md)

