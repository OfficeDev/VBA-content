---
title: TextInput.Clear Method (Word)
keywords: vbawd10.chm153550949
f1_keywords:
- vbawd10.chm153550949
ms.prod: word
api_name:
- Word.TextInput.Clear
ms.assetid: 863fc6e4-efb6-3d3a-5f4f-19caab70f44f
ms.date: 06/08/2017
---


# TextInput.Clear Method (Word)

Deletes the text from the specified text form field.


## Syntax

 _expression_ . **Clear**

 _expression_ Required. A variable that represents a **[TextInput](textinput-object-word.md)** object.


## Example

This example protects the document for forms and deletes the text from the first form field if the field is a text form field.


```vb
ActiveDocument.Protect Type:=wdAllowOnlyFormFields, NoReset:=True 
If ActiveDocument.FormFields(1).Type = wdFieldFormTextInput Then 
 ActiveDocument.FormFields(1).TextInput.Clear 
End If
```


## See also


#### Concepts


[TextInput Object](textinput-object-word.md)

