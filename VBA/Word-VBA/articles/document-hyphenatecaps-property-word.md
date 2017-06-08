---
title: Document.HyphenateCaps Property (Word)
keywords: vbawd10.chm158007308
f1_keywords:
- vbawd10.chm158007308
ms.prod: word
api_name:
- Word.Document.HyphenateCaps
ms.assetid: 13f421aa-7e37-4f13-9b34-7ed139421e17
ms.date: 06/08/2017
---


# Document.HyphenateCaps Property (Word)

 **True** if words in all capital letters can be hyphenated. Read/write **Boolean** .


## Syntax

 _expression_ . **HyphenateCaps**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example enables automatic hyphenation for the active document and allows capitalized words to be hyphenated.


```vb
With ActiveDocument 
 .AutoHyphenation = True 
 .HyphenateCaps = True 
End With
```


## See also


#### Concepts


[Document Object](document-object-word.md)

