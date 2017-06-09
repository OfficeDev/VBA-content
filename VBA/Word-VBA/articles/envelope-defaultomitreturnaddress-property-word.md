---
title: Envelope.DefaultOmitReturnAddress Property (Word)
keywords: vbawd10.chm152567817
f1_keywords:
- vbawd10.chm152567817
ms.prod: word
api_name:
- Word.Envelope.DefaultOmitReturnAddress
ms.assetid: d1ef3e8d-4410-61b4-0631-6d458dcb14b8
ms.date: 06/08/2017
---


# Envelope.DefaultOmitReturnAddress Property (Word)

 **True** if the return address is omitted from envelopes by default. Read/write **Boolean** .


## Syntax

 _expression_ . **DefaultOmitReturnAddress**

 _expression_ A variable that represents a **[Envelope](envelope-object-word.md)** object.


## Example

This example omits return addresses from new envelopes by default.


```vb
ActiveDocument.Envelope.DefaultOmitReturnAddress = True
```

This example displays the return address status in a message box.




```vb
If ActiveDocument.Envelope.DefaultOmitReturnAddress = True Then 
 MsgBox "A return address is not included by default." 
Else 
 MsgBox "A return address is included by default." 
End If
```


## See also


#### Concepts


[Envelope Object](envelope-object-word.md)

