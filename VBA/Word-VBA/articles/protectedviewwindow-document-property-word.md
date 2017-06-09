---
title: ProtectedViewWindow.Document Property (Word)
keywords: vbawd10.chm231735297
f1_keywords:
- vbawd10.chm231735297
ms.prod: word
api_name:
- Word.ProtectedViewWindow.Document
ms.assetid: a4a3e32e-a697-9d9a-f4ea-a07daa1ea238
ms.date: 06/08/2017
---


# ProtectedViewWindow.Document Property (Word)

Returns a [Document](document-object-word.md) object associated with the protected view window. Read-only.


## Syntax

 _expression_ . **Document**

 _expression_ A variable that represents a **[ProtectedViewWindow](protectedviewwindow-object-word.md)** object.


## Remarks

A document displayed in a protected view window is not a member of the  **[Documents](application-documents-property-word.md)** collection. Instead, use the **Document** property to access a document that is displayed in a protected view window.


## Example

The following code example displays the name of the document in the active protected view window.


```vb
Dim myDoc As Document 
 
Set myDoc = ActiveProtectedViewWindow.Document 
MsgBox myDoc.Name
```


## See also


#### Concepts


[ProtectedViewWindow Object](protectedviewwindow-object-word.md)

