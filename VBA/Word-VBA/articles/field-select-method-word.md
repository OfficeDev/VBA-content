---
title: Field.Select Method (Word)
keywords: vbawd10.chm154140671
f1_keywords:
- vbawd10.chm154140671
ms.prod: word
api_name:
- Word.Field.Select
ms.assetid: 03fa304c-acc7-30a5-7dfa-06098bbdac7a
ms.date: 06/08/2017
---


# Field.Select Method (Word)

Selects the specified field.


## Syntax

 _expression_ . **Select**

 _expression_ Required. A variable that represents a **[Field](field-object-word.md)** object.


## Remarks

After using this method, use the  **[Selection](selection-object-word.md)** object to work with the selected items. For more information, see[Working with the Selection Object](http://msdn.microsoft.com/library/a1ef7e48-5a0f-d278-4b67-7b96f4e24052%28Office.15%29.aspx).


## Example

This example updates and selects the first field in the active document.


```vb
ActiveDocument.ActiveWindow.View.FieldShading = _ 
 wdFieldShadingWhenSelected 
If ActiveDocument.Fields.Count >= 1 Then 
 With ActiveDocument.Fields(1) 
 .Update 
 .Select 
 End With 
End If
```


## See also


#### Concepts


[Field Object](field-object-word.md)

