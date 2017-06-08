---
title: Envelope.ReturnAddress Property (Word)
keywords: vbawd10.chm152567810
f1_keywords:
- vbawd10.chm152567810
ms.prod: word
api_name:
- Word.Envelope.ReturnAddress
ms.assetid: cbbbcc74-afb9-f646-caf8-171605de48c8
ms.date: 06/08/2017
---


# Envelope.ReturnAddress Property (Word)

Returns a  **Range** object that represents the envelope return address.


## Syntax

 _expression_ . **ReturnAddress**

 _expression_ Required. A variable that represents an **[Envelope](envelope-object-word.md)** object.


## Remarks

An error occurs if you use this property before adding an envelope to the document.


## Example

This example displays the return address if an envelope has been added to the active document; otherwise, a message box is displayed.


```vb
On Error GoTo errhandler 
addr = ActiveDocument.Envelope.ReturnAddress.Text 
MsgBox Prompt:=addr, Title:="Return Address" 
errhandler: 
If Err = 5852 Then MsgBox _ 
 "The active document doesn't include an envelope"
```


## See also


#### Concepts


[Envelope Object](envelope-object-word.md)

