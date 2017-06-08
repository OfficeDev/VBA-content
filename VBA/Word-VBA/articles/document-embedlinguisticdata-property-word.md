---
title: Document.EmbedLinguisticData Property (Word)
keywords: vbawd10.chm158007673
f1_keywords:
- vbawd10.chm158007673
ms.prod: word
api_name:
- Word.Document.EmbedLinguisticData
ms.assetid: ad76bcba-dad3-6745-8cdb-a56797054af4
ms.date: 06/08/2017
---


# Document.EmbedLinguisticData Property (Word)

 **True** for Microsoft Word to embed speech and handwriting so that data can be converted back to speech or handwriting. Read/write **Boolean** .


## Syntax

 _expression_ . **EmbedLinguisticData**

 _expression_ An expression that returns a **[Document](document-object-word.md)** object.


## Remarks

 The **EmbedLinguisticData** property also allows you to store East Asian IME keystrokes to improve correction and controls text service data received from devices connected to Microsoft Office using the Windows Text Service Framework Application Programming Interface


## Example

This example embeds into the active document any speech or handwriting that may exist in the document.


```vb
Sub EmbedSpeechHandwriting() 
 ActiveDocument.EmbedLinguisticData = True 
End Sub
```


## See also


#### Concepts


[Document Object](document-object-word.md)

