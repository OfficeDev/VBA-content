---
title: Fields.Update Method (Word)
keywords: vbawd10.chm154140773
f1_keywords:
- vbawd10.chm154140773
ms.prod: word
api_name:
- Word.Fields.Update
ms.assetid: 55aaae86-015f-fc4f-ff7c-42fddad05c27
ms.date: 06/08/2017
---


# Fields.Update Method (Word)

Updates the result of the fields object.


## Syntax

 _expression_ . **Update**

 _expression_ Required. A variable that represents a **[Fields](fields-object-word.md)** collection.


### Return Value

Long


## Remarks

Returns 0 (zero) if no errors occur when the fields are updated, or returns a  **Long** that represents the index of the first field that contains an error.


## Example

This example updates all the fields in the main story (that is, the main body) of the active document. A return value of 0 (zero) indicates that the fields were updated without error.


```vb
If ActiveDocument.Fields.Update = 0 Then 
 MsgBox "Update Successful" 
Else 
 MsgBox "Field " &; ActiveDocument.Fields.Update &; _ 
 " has an error" 
End If
```


## See also


#### Concepts


[Fields Collection Object](fields-object-word.md)

