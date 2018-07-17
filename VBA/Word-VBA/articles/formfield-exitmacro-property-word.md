---
title: FormField.ExitMacro Property (Word)
keywords: vbawd10.chm153616388
f1_keywords:
- vbawd10.chm153616388
ms.prod: word
api_name:
- Word.FormField.ExitMacro
ms.assetid: b8930661-e02f-e058-571e-986da33a477d
ms.date: 06/08/2017
---


# FormField.ExitMacro Property (Word)

Returns or sets an exit macro name for the specified form field (CheckBox, DropDown, or TextInput). Read/write  **String** .


## Syntax

 _expression_ . **ExitMacro**

 _expression_ A variable that represents a **[FormField](formfield-object-word.md)** object.


## Remarks

The exit macro runs when the form field loses the focus. 


## Example

This example assigns the macro named "Reformat" to the first form field in the selection.


```vb
If Selection.FormFields.Count > 0 Then _ 
 Selection.FormFields(1).ExitMacro = "Reformat"
```

This example assigns the macro named "Blue" to the last form field in "Form.doc."




```vb
Dim intMax As Integer 
 
intMax = Documents("Form.doc").FormFields.Count 
Documents("Form.doc").FormFields(intMax).ExitMacro = "Blue"
```


## See also


#### Concepts


[FormField Object](formfield-object-word.md)

