---
title: Envelope.DefaultFaceUp Property (Word)
keywords: vbawd10.chm152567828
f1_keywords:
- vbawd10.chm152567828
ms.prod: word
api_name:
- Word.Envelope.DefaultFaceUp
ms.assetid: ce745551-4385-420d-1790-464bf03da3d9
ms.date: 06/08/2017
---


# Envelope.DefaultFaceUp Property (Word)

 **True** if envelopes are fed face up by default. Read/write **Boolean** .


## Syntax

 _expression_ . **DefaultFaceUp**

 _expression_ A variable that represents a **[Envelope](envelope-object-word.md)** object.


## Example

This example sets envelopes to be fed face up by default. The UpdateDocument method updates the envelope in the active document.


```vb
With ActiveDocument.Envelope 
 .DefaultFaceUp = True 
 .DefaultOrientation = wdCenterPortrait 
 .UpdateDocument 
End With
```

This example displays a message telling the user how to feed the envelopes into the printer based on the default setting.




```vb
If ActiveDocument.Envelope.DefaultFaceUp = True Then 
 MsgBox "Feed envelopes face up." 
Else 
 MsgBox "Feed envelopes face down." 
End If
```


## See also


#### Concepts


[Envelope Object](envelope-object-word.md)

