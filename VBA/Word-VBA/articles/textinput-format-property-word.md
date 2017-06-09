---
title: TextInput.Format Property (Word)
keywords: vbawd10.chm153550851
f1_keywords:
- vbawd10.chm153550851
ms.prod: word
api_name:
- Word.TextInput.Format
ms.assetid: 5950cabb-dfd3-0107-6a51-efe8813a297f
ms.date: 06/08/2017
---


# TextInput.Format Property (Word)

Returns the text formatting for the specified text box. Read-only  **String** .


## Syntax

 _expression_ . **Format**

 _expression_ Required. A variable that represents a **[TextInput](textinput-object-word.md)** object.


## Example

This example displays the text formatting in the first field of the active document.


```vb
If ActiveDocument.FormFields(1).Type = wdFieldFormTextInput Then 
 MsgBox ActiveDocument.FormFields(1).TextInput.Format 
Else 
 MsgBox "First field is not a text form field" 
End If
```


## See also


#### Concepts


[TextInput Object](textinput-object-word.md)

