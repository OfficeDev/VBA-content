---
title: Document.WebOptions Property (Word)
keywords: vbawd10.chm158007626
f1_keywords:
- vbawd10.chm158007626
ms.prod: word
api_name:
- Word.Document.WebOptions
ms.assetid: 038eef42-8c57-8910-d8c1-7b9937f180c5
ms.date: 06/08/2017
---


# Document.WebOptions Property (Word)

Returns the  **[WebOptions](weboptions-object-word.md)** object, which contains document-level attributes used by Microsoft Word when you save a document as a Web page or open a Web page. Read-only.


## Syntax

 _expression_ . **WebOptions**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example specifies that cascading style sheets and Western document encoding be used when items in the active document are saved to a Web page.


```vb
Set objWO = ActiveDocument.WebOptions 
objWO.RelyOnCSS = True 
objWO.Encoding = msoEncodingWestern
```


## See also


#### Concepts


[Document Object](document-object-word.md)

