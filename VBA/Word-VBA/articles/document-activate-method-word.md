---
title: Document.Activate Method (Word)
keywords: vbawd10.chm158007409
f1_keywords:
- vbawd10.chm158007409
ms.prod: word
api_name:
- Word.Document.Activate
ms.assetid: 83cc5935-020b-470a-f7aa-7fea057ec08b
ms.date: 06/08/2017
---


# Document.Activate Method (Word)

Activates the specified document so that it becomes the active document.


## Syntax

 _expression_ . **Activate**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example activates the document named "Sales.doc."


```vb
Sub OpenSales() 
 'Sales.doc must exist and be open but not active. 
 Documents("Sales.doc").Activate 
End Sub
```


## See also


#### Concepts


[Document Object](document-object-word.md)

