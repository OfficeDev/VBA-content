---
title: FormField.Previous Property (Word)
keywords: vbawd10.chm153616399
f1_keywords:
- vbawd10.chm153616399
ms.prod: word
api_name:
- Word.FormField.Previous
ms.assetid: 34e8d20a-5009-67eb-fdc0-bafad134e9b3
ms.date: 06/08/2017
---


# FormField.Previous Property (Word)

Returns the previous form field in the collection. Read-only.


## Syntax

 _expression_ . **Previous**

 _expression_ A variable that represents a **[FormField](formfield-object-word.md)** object.


## Example

This example displays the field code of the second-to-last form field in the active document.


```vb
Set aField = ActiveDocument _ 
 .FormFields(ActiveDocument.FormFields.Count).Previous 
MsgBox "Form field code = " &; aField.Code
```


## See also


#### Concepts


[FormField Object](formfield-object-word.md)

