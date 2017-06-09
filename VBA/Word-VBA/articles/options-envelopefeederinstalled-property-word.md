---
title: Options.EnvelopeFeederInstalled Property (Word)
keywords: vbawd10.chm162988067
f1_keywords:
- vbawd10.chm162988067
ms.prod: word
api_name:
- Word.Options.EnvelopeFeederInstalled
ms.assetid: 9b614965-d1e2-21df-a6f5-b595d48c6227
ms.date: 06/08/2017
---


# Options.EnvelopeFeederInstalled Property (Word)

 **True** if the current printer has a special feeder for envelopes. Read-only **Boolean** .


## Syntax

 _expression_ . **EnvelopeFeederInstalled**

 _expression_ A variable that represents a **[Options](options-object-word.md)** object.


## Example

This example prints the active document as an envelope, provided that there is an envelope feeder installed.


```vb
If Options.EnvelopeFeederInstalled = True Then 
 ActiveDocument.Envelope.PrintOut _ 
 AddressFromLeft:=InchesToPoints(3), _ 
 AddressFromTop:=InchesToPoints(1.5) 
Else 
 Msgbox "No envelope feeder available." 
End If
```


## See also


#### Concepts


[Options Object](options-object-word.md)

