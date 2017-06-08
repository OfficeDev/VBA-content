---
title: FormField.Enabled Property (Word)
keywords: vbawd10.chm153616393
f1_keywords:
- vbawd10.chm153616393
ms.prod: word
api_name:
- Word.FormField.Enabled
ms.assetid: 1002dfdd-387e-9c44-27aa-c855e78784bc
ms.date: 06/08/2017
---


# FormField.Enabled Property (Word)

 **True** if a form field is enabled. Read/write **Boolean** .


## Syntax

 _expression_ . **Enabled**

 _expression_ An expression that represents a **[FormField](formfield-object-word.md)** object.


## Remarks

If a form field is enabled, its contents can be changed as the form is filled in.


## Example

If the first form field in the active document is an enabled check box, this example selects the check box.


```vb
Dim ffFirst As FormField 
 
Set ffFirst = ActiveDocument.FormFields(1) 
If ffFirst.Enabled = True And _ 
 ffFirst.Type = wdFieldFormCheckBox Then 
 ffFirst.CheckBox.Value = True 
End If
```


## See also


#### Concepts


[FormField Object](formfield-object-word.md)

