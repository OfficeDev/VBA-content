---
title: Envelope.Vertical Property (Word)
keywords: vbawd10.chm152567830
f1_keywords:
- vbawd10.chm152567830
ms.prod: word
api_name:
- Word.Envelope.Vertical
ms.assetid: 23f8fbf0-375e-98c2-81b4-451cc8973e85
ms.date: 06/08/2017
---


# Envelope.Vertical Property (Word)

 **True** vertically orients text on Asian envelopes. Read/write **Boolean** .


## Syntax

 _expression_ . **Vertical**

 _expression_ Required. A variable that represents an **[Envelope](envelope-object-word.md)** object.


## Example

This example determines if the active document is a mail merge envelope document and if the language setting is Chinese, and if so, sets the envelope's orientation to vertical and updates the current document.


```vb
Sub VerticalEnvelope() 
 If ActiveDocument.MailMerge.MainDocumentType = wdEnvelopes And 
 Application.Language = msoLanguageIDChineseHongKong Then 
 With ActiveDocument.Envelope 
 .Vertical = True 
 .UpdateDocument 
 End With 
 End If 
End Sub
```


## See also


#### Concepts


[Envelope Object](envelope-object-word.md)

