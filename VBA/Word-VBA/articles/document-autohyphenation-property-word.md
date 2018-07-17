---
title: Document.AutoHyphenation Property (Word)
keywords: vbawd10.chm158007307
f1_keywords:
- vbawd10.chm158007307
ms.prod: word
api_name:
- Word.Document.AutoHyphenation
ms.assetid: 17e53212-3717-c8a1-7f39-464622a6cd65
ms.date: 06/08/2017
---


# Document.AutoHyphenation Property (Word)

 **True** if automatic hyphenation is turned on for the specified document. Read/write **Boolean** .


## Syntax

 _expression_ . **AutoHyphenation**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Example

This example turns on automatic hyphenation, with a hyphenation zone of 0.25 inch. Words in all capital letters aren't hyphenated.


```vb
With ActiveDocument 
 .HyphenationZone = InchesToPoints(0.25) 
 .HyphenateCaps = False 
 .AutoHyphenation = True 
End With
```


## See also


#### Concepts


[Document Object](document-object-word.md)

