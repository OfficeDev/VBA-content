---
title: Fields.Locked Property (Word)
keywords: vbawd10.chm154140674
f1_keywords:
- vbawd10.chm154140674
ms.prod: word
api_name:
- Word.Fields.Locked
ms.assetid: 9ecebbac-fc22-0474-ed2e-a17a549d6722
ms.date: 06/08/2017
---


# Fields.Locked Property (Word)

 **True** if all fields in the **Fields** collection are locked. Read/write **Long** .


## Syntax

 _expression_ . **Locked**

 _expression_ Required. A variable that represents a **[Fields](fields-object-word.md)** collection.


## Remarks

This property can be  **True** , **False** , or **wdUndefined** (if some of the fields in the collection are locked and others are not).


## Example

This example locks all the fields in the selection.


```vb
Selection.Fields.Locked = True
```

This example displays a message if some of the fields in the active document are locked.




```vb
Set theFields = ActiveDocument.Fields 
If theFields.Locked = wdUndefined Then 
 MsgBox "Some fields are locked" 
ElseIf theFields.Locked = False Then 
 MsgBox "No fields are locked" 
ElseIf theFields.Locked = True Then 
 MsgBox "All fields are locked" 
End If
```


## See also


#### Concepts


[Fields Collection Object](fields-object-word.md)

