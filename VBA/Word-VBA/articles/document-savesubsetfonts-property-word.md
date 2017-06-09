---
title: Document.SaveSubsetFonts Property (Word)
keywords: vbawd10.chm158007349
f1_keywords:
- vbawd10.chm158007349
ms.prod: word
api_name:
- Word.Document.SaveSubsetFonts
ms.assetid: 01210b29-f346-e513-6876-3dab30b940e1
ms.date: 06/08/2017
---


# Document.SaveSubsetFonts Property (Word)

 **True** if Microsoft Word saves a subset of the embedded TrueType fonts with the document. Read/write **Boolean** .


## Syntax

 _expression_ . **SaveSubsetFonts**

 _expression_ A variable that represents a **[Document](document-object-word.md)** object.


## Remarks

If fewer than 32 characters of a TrueType font are used in a document, Word embeds the subset (only the characters used) in the document. If more than 32 characters are used, Word embeds the entire font.


## Example

This example sets a document named "MyDoc" to save only a subset of its embedded TrueType fonts (when just a few characters are used), and then it saves "MyDoc."


```vb
With Documents("MyDoc") 
 .EmbedTrueTypeFonts = True 
 .SaveSubsetFonts = True 
 .Save 
End With
```


## See also


#### Concepts


[Document Object](document-object-word.md)

