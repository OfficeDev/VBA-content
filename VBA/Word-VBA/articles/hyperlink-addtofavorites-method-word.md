---
title: Hyperlink.AddToFavorites Method (Word)
keywords: vbawd10.chm161284201
f1_keywords:
- vbawd10.chm161284201
ms.prod: word
api_name:
- Word.Hyperlink.AddToFavorites
ms.assetid: 262f05e9-3697-a695-db2d-39162158ec41
ms.date: 06/08/2017
---


# Hyperlink.AddToFavorites Method (Word)

Creates a shortcut to the document or hyperlink and adds it to the Favorites folder.


## Syntax

 _expression_ . **AddToFavorites**

 _expression_ Required. A variable that represents a **[Hyperlink](hyperlink-object-word.md)** object.


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


[Hyperlink Object](hyperlink-object-word.md)

