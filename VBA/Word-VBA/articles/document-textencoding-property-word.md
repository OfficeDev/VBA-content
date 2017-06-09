---
title: Document.TextEncoding Property (Word)
keywords: vbawd10.chm158007653
f1_keywords:
- vbawd10.chm158007653
ms.prod: word
api_name:
- Word.Document.TextEncoding
ms.assetid: a11b45c1-1829-0df0-3403-e92268d9ec81
ms.date: 06/08/2017
---


# Document.TextEncoding Property (Word)

Returns or sets the code page, or character set, that Microsoft Word uses for a document saved as an encoded text file. Read/write  **MsoEncoding** .


## Syntax

 _expression_ . **TextEncoding**

 _expression_ Required. A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

The  **TextEncoding** property sets text encoding separately from HTML encoding, which you can set using the **Encoding** property. To set text encoding for all documents saved as text files, use the **DefaultTextEncoding** property.


## Example

This example sets the text encoding for the active document to Japanese if it is saved as a text file.


```vb
Sub EncodeText() 
 ActiveDocument.TextEncoding = msoEncodingJapaneseShiftJIS 
End Sub
```


## See also


#### Concepts


[Document Object](document-object-word.md)

