---
title: Field.Update Method (Word)
keywords: vbawd10.chm154075237
f1_keywords:
- vbawd10.chm154075237
ms.prod: word
api_name:
- Word.Field.Update
ms.assetid: e4e941aa-3223-ae0b-8366-9e14d92fff52
ms.date: 06/08/2017
---


# Field.Update Method (Word)

Updates the result of the field. Returns  **True** if the field is updated successfully.


## Syntax

 _expression_ . **Update**

 _expression_ Required. A variable that represents a **[Field](field-object-word.md)** object.


### Return Value

Boolean


## Example

This example updates the first field in the active document. A return value of 1 (True) indicates that the fields were updated without error.


```vb
If ActiveDocument.Fields(0).Update = 1 Then 
 MsgBox "Update Successful" 
Else 
 MsgBox "Field " &; ActiveDocument.Fields(0).Update &; _ 
 " has an error" 
End If
```


## See also


#### Concepts


[Field Object](field-object-word.md)

