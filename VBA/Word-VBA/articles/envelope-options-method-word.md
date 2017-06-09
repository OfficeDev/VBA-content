---
title: Envelope.Options Method (Word)
keywords: vbawd10.chm152567912
f1_keywords:
- vbawd10.chm152567912
ms.prod: word
api_name:
- Word.Envelope.Options
ms.assetid: 5619bf1a-eaf9-aa0e-01c3-66111c20dd0c
ms.date: 06/08/2017
---


# Envelope.Options Method (Word)

Displays the  **Envelope Options** dialog box.


## Syntax

 _expression_ . **Options**

 _expression_ Required. A variable that represents an **[Envelope](envelope-object-word.md)** object.


## Remarks

The  **Options** method works only if the document is the main document of an envelope mail merge.


## Example

This example checks that the active document is an envelope mail merge main document, and if it is, displays the  **Envelope Options** dialog box.


```vb
Sub EnvelopeOptions() 
 If ActiveDocument.MailMerge.MainDocumentType = wdEnvelopes Then 
 ActiveDocument.Envelope.Options 
 End If 
End Sub
```


## See also


#### Concepts


[Envelope Object](envelope-object-word.md)

