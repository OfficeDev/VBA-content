---
title: Envelope.DefaultPrintFIMA Property (Word)
keywords: vbawd10.chm152567813
f1_keywords:
- vbawd10.chm152567813
ms.prod: word
api_name:
- Word.Envelope.DefaultPrintFIMA
ms.assetid: 13cba63f-dc2a-722e-1bc2-21db8c0e82cd
ms.date: 06/08/2017
---


# Envelope.DefaultPrintFIMA Property (Word)

 **True** to add a Facing Identification Mark (FIM-A) to envelopes by default. Read/write **Boolean** .


## Syntax

 _expression_ . **DefaultPrintFIMA**

 _expression_ A variable that represents a **[Envelope](envelope-object-word.md)** object.


## Remarks

This property applies to U.S. mail only. A FIM-A code is used to presort courtesy reply mail.


## Example

This example sets the default envelope settings to include a Facing Identification Mark (FIM-A).


```vb
With ActiveDocument.Envelope 
 .DefaultPrintFIMA = True 
End With
```


## See also


#### Concepts


[Envelope Object](envelope-object-word.md)

