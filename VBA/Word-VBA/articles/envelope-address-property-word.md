---
title: Envelope.Address Property (Word)
keywords: vbawd10.chm152567809
f1_keywords:
- vbawd10.chm152567809
ms.prod: word
api_name:
- Word.Envelope.Address
ms.assetid: 01d6d211-a4f1-c3cd-470c-f49d6bb22fe6
ms.date: 06/08/2017
---


# Envelope.Address Property (Word)

Returns the envelope delivery address as a  **Range** object. Read-only.


## Syntax

 _expression_ . **Address**

 _expression_ Required. A variable that represents an **[Envelope](envelope-object-word.md)** object.


## Example

This example displays the delivery address if an envelope has been added to the document; otherwise, it displays a message.


```vb
On Error GoTo errhandler 
addr = ActiveDocument.Envelope.Address.Text 
MsgBox Prompt:=addr, Title:="Delivery Address" 
errhandler: 
If Err = 5852 Then MsgBox "Insert an envelope into the document"
```


## See also


#### Concepts


[Envelope Object](envelope-object-word.md)

