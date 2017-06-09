---
title: EmailSignature.NewMessageSignature Property (Word)
keywords: vbawd10.chm165412967
f1_keywords:
- vbawd10.chm165412967
ms.prod: word
api_name:
- Word.EmailSignature.NewMessageSignature
ms.assetid: fed9f151-47b8-3e76-1764-b6e80bdbfb5e
ms.date: 06/08/2017
---


# EmailSignature.NewMessageSignature Property (Word)

Returns or sets the signature that Microsoft Word appends to new e-mail messages. Read/write  **String** .


## Syntax

 _expression_ . **NewMessageSignature**

 _expression_ An expression that returns an **[EmailSignature](emailsignature-object-word.md)** object.


## Remarks

When setting this property, you must use the name of an e-mail signature that you have created in the  **E-mail Options** dialog box, available from the **General** tab of the **Options** dialog box ( **Tools** menu).


## Example

This example changes the signature Word appends to new outgoing e-mail messages.


```vb
With Application.EmailOptions.EmailSignature 
 .NewMessageSignature = "Signature1" 
End With
```


## See also


#### Concepts


[EmailSignature Object](emailsignature-object-word.md)

