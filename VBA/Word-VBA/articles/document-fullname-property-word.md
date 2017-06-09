---
title: Document.FullName Property (Word)
keywords: vbawd10.chm158007325
f1_keywords:
- vbawd10.chm158007325
ms.prod: word
api_name:
- Word.Document.FullName
ms.assetid: 795a20cb-c744-6c3c-8e7f-f7a749489819
ms.date: 06/08/2017
---


# Document.FullName Property (Word)

Returns a  **String** that represents the name of a document, including the path. Read-only.


## Syntax

 _expression_ . **FullName**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

Using this property is equivalent to using the  **Path** , **PathSeparator** , and **Name** properties in sequence.


## Example

This example displays the path and file name of the active document.


```vb
Sub DocName() 
 MsgBox ActiveDocument.FullName 
End Sub
```


## See also


#### Concepts


[Document Object](document-object-word.md)

