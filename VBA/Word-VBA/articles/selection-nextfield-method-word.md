---
title: Selection.NextField Method (Word)
keywords: vbawd10.chm158662834
f1_keywords:
- vbawd10.chm158662834
ms.prod: word
api_name:
- Word.Selection.NextField
ms.assetid: 40007462-3bb5-59a7-89cb-27d654795e76
ms.date: 06/08/2017
---


# Selection.NextField Method (Word)

Selects the next field.


## Syntax

 _expression_ . **NextField**

 _expression_ Required. A variable that represents a **[Selection](selection-object-word.md)** object.


### Return Value

Field


## Remarks

If this method finds a field, it returns a  **Field** object; if not, it returns **Nothing** .


## Example

This example updates the next field in the selection.


```vb
If Not (Selection.NextField Is Nothing) Then 
 Selection.Fields.Update 
End If
```

This example selects the next field in the selection, and if a field is found, displays a message in the status bar.




```vb
Set myField = Selection.NextField 
If Not (myField Is Nothing) Then StatusBar = "Field found"
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

