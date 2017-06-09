---
title: Document.OpenEncoding Property (Word)
keywords: vbawd10.chm158007628
f1_keywords:
- vbawd10.chm158007628
ms.prod: word
api_name:
- Word.Document.OpenEncoding
ms.assetid: a147f531-de42-47c5-1a74-12ea65e64b8b
ms.date: 06/08/2017
---


# Document.OpenEncoding Property (Word)

Returns the encoding used to open the specified document. Read-only  **MsoEncoding** .


## Syntax

 _expression_ . **OpenEncoding**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example tests whether the current document was opened with UTF7 encoding.


```vb
If ActiveDocument.OpenEncoding = msoEncodingUTF7 Then 
 MsgBox "This is a UTF7-encoded text file!" 
Else 
 MsgBox "This is not a UTF7-encoded text file!" 
End If
```


## See also


#### Concepts


[Document Object](document-object-word.md)

