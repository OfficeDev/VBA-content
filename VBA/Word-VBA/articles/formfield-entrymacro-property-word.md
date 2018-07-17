---
title: FormField.EntryMacro Property (Word)
keywords: vbawd10.chm153616387
f1_keywords:
- vbawd10.chm153616387
ms.prod: word
api_name:
- Word.FormField.EntryMacro
ms.assetid: db4ff78e-6795-0e8e-20db-56ceac01b8f2
ms.date: 06/08/2017
---


# FormField.EntryMacro Property (Word)

Returns or sets an entry macro name for the specified form field (CheckBox, DropDown, or TextInput). Read/write  **String** .


## Syntax

 _expression_ . **EntryMacro**

 _expression_ A variable that represents a **[FormField](formfield-object-word.md)** object.


## Remarks

The entry macro runs when the form field gets the focus. 


## Example

This example assigns the macro named "Blue" to the first form field in "Form.doc."


```
Documents("Form.doc").FormFields(1).EntryMacro = "Blue"
```

This example assigns the macro named "Breadth" to the form field named "Text1" in the active document.




```vb
ActiveDocument.FormFields("Text1").EntryMacro = "Breadth"
```


## See also


#### Concepts


[FormField Object](formfield-object-word.md)

