---
title: EmailOptions.EmailSignature Property (Word)
keywords: vbawd10.chm165347436
f1_keywords:
- vbawd10.chm165347436
ms.prod: word
api_name:
- Word.EmailOptions.EmailSignature
ms.assetid: 853e0b8d-8e25-4626-154f-1d634e485929
ms.date: 06/08/2017
---


# EmailOptions.EmailSignature Property (Word)

Returns an  **[EmailSignature](emailsignature-object-word.md)** object that represents the signatures Microsoft Word appends to outgoing e-mail messages. Read-only.


## Syntax

 _expression_ . **EmailSignature**

 _expression_ A variable that represents a **[EmailOptions](emailoptions-object-word.md)** object.


## Example

This example displays the signature Word appends to new outgoing e-mail messages.


```vb
With Application.EmailOptions.EmailSignature 
 If .NewMessageSignature = "" Then 
 MsgBox "There is no signature for new " _ 
 &; "e-mail messages!" 
 Else 
 MsgBox "The signature for new e-mail" _ 
 &; "messages is: " &; vbLf &; vbLf _ 
 &; .NewMessageSignature 
 End If 
End With
```


## See also


#### Concepts


[EmailOptions Object](emailoptions-object-word.md)

