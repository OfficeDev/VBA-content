---
title: CheckBox.Valid Property (Word)
keywords: vbawd10.chm153485312
f1_keywords:
- vbawd10.chm153485312
ms.prod: word
api_name:
- Word.CheckBox.Valid
ms.assetid: 5f14faf3-8025-709d-67a4-7ba0ae46b467
ms.date: 06/08/2017
---


# CheckBox.Valid Property (Word)

 **True** if the specified form field object is a valid check box form field. Read-only **Boolean** .


## Syntax

 _expression_ . **Valid**

 _expression_ A variable that represents a **[CheckBox](checkbox-object-word.md)** object.


## Example

This example adds a text form field at the insertion point. Because  `myFormField` is a text input field and not a check box, the message box displays "False."


```vb
Selection.Collapse Direction:=wdCollapseStart 
Set myFormField = ActiveDocument.FormFields.Add(Range:= _ 
 Selection.Range, Type:=wdFieldFormTextInput) 
MsgBox myFormField.CheckBox.Valid
```


## See also


#### Concepts


[CheckBox Object](checkbox-object-word.md)

