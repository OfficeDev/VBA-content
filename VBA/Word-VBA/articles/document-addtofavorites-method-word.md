---
title: Document.AddToFavorites Method (Word)
keywords: vbawd10.chm158007432
f1_keywords:
- vbawd10.chm158007432
ms.prod: word
api_name:
- Word.Document.AddToFavorites
ms.assetid: e810df76-18a8-d6b8-8d72-fb6386e6ce3a
ms.date: 06/08/2017
---


# Document.AddToFavorites Method (Word)

Creates a shortcut to the document or hyperlink and adds it to the Favorites folder.


## Syntax

 _expression_ . **AddToFavorites**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example creates a shortcut for each hyperlink in the active document and adds it to the Favorites folder.


```vb
For Each myHyperlink In ActiveDocument.Hyperlinks 
 myHyperlink.AddToFavorites 
Next myHyperlink
```

This example creates a shortcut to Sales.doc and adds it to the Favorites folder. If Sales.doc isn't currently open, this example opens it from the C:\Documents folder.




```vb
For Each doc in Documents 
 If LCase(doc.Name) = "sales.doc" Then isOpen = True 
Next doc 
If isOpen <> True Then Documents.Open _ 
 FileName:="C:\Documents\Sales.doc" 
Documents("Sales.doc").AddToFavorites
```


## See also


#### Concepts


[Document Object](document-object-word.md)

