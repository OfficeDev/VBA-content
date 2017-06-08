---
title: Document.DoNotEmbedSystemFonts Property (Word)
keywords: vbawd10.chm158007634
f1_keywords:
- vbawd10.chm158007634
ms.prod: word
api_name:
- Word.Document.DoNotEmbedSystemFonts
ms.assetid: 435054c0-f7e3-e206-146d-7e29cce2c71d
ms.date: 06/08/2017
---


# Document.DoNotEmbedSystemFonts Property (Word)

 **True** for Microsoft Word to not embed common system fonts. Read/write **Boolean** .


## Syntax

 _expression_ . **DoNotEmbedSystemFonts**

 _expression_ An expression that returns a **[Document](document-object-word.md)** object.


## Remarks

Setting the  **[Document](document-object-word.md)** property to **False** is useful if the user is on an East Asian system and wants to create a document that is readable by others who do not have fonts for that language on their system. For example, a user on a Japanese system could choose to embed the fonts in a document so that the Japanese document would be readable on all systems.


## Example

This example embeds all fonts in the current document.


```vb
Sub EmbedFonts() 
 With ActiveDocument 
 If .EmbedTrueTypeFonts = False Then 
 .EmbedTrueTypeFonts = True 
 .DoNotEmbedSystemFonts = False 
 Else 
 .DoNotEmbedSystemFonts = False 
 End If 
 End With 
End Sub
```


## See also


#### Concepts


[Document Object](document-object-word.md)

