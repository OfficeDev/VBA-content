---
title: Envelope.SenderNamefromLeft Property (Word)
keywords: vbawd10.chm152567835
f1_keywords:
- vbawd10.chm152567835
ms.prod: word
api_name:
- Word.Envelope.SenderNamefromLeft
ms.assetid: 3cd5009a-d500-3346-ffea-957f905d3929
ms.date: 06/08/2017
---


# Envelope.SenderNamefromLeft Property (Word)

Returns or sets a  **Single** that represents the position, measured in points, of the sender's name from the left edge of the envelope. Read/write.


## Syntax

 _expression_ . **SenderNamefromLeft**

 _expression_ An expression that returns an **[Envelope](envelope-object-word.md)** object.


## Remarks

Use this property for Asian language envelopes.


## Example

This example checks that the active document is a mail merge envelope and that it is formatted for vertical type. If so, it positions the recipient and sender address information.


```vb
Sub NewEnvelopeMerge() 
 With ActiveDocument 
 If .MailMerge.MainDocumentType = wdEnvelopes Then 
 With ActiveDocument.Envelope 
 If .Vertical = True Then 
 .RecipientNamefromLeft = InchesToPoints(2.5) 
 .RecipientNamefromTop = InchesToPoints(2) 
 .RecipientPostalfromLeft = InchesToPoints(1.5) 
 .RecipientPostalfromTop = InchesToPoints(0.5) 
 .SenderNamefromLeft = InchesToPoints(0.5) 
 .SenderNamefromTop = InchesToPoints(2) 
 .SenderPostalfromLeft = InchesToPoints(0.5) 
 .SenderPostalfromTop = InchesToPoints(3) 
 End If 
 End With 
 End If 
 End With 
End Sub
```


## See also


#### Concepts


[Envelope Object](envelope-object-word.md)

