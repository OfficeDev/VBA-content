---
title: Field.Locked Property (Word)
keywords: vbawd10.chm154075138
f1_keywords:
- vbawd10.chm154075138
ms.prod: word
api_name:
- Word.Field.Locked
ms.assetid: 2f1b1351-8de1-f2b0-0c39-b944bf23a92e
ms.date: 06/08/2017
---


# Field.Locked Property (Word)

 **True** if the specified field is locked. Read/write **Boolean** .


## Syntax

 _expression_ . **Locked**

 _expression_ Required. A variable that represents a **[Field](field-object-word.md)** object.


## Remarks

When a field is locked, you cannot update the field results.


## Example

This example inserts a DATE field at the beginning of the selection and then locks the field.


```vb
Selection.Collapse Direction:=wdCollapseStart 
Set myField = ActiveDocument.Fields.Add(Range:=Selection.Range, _ 
 Type:=wdFieldDate) 
myField.Locked = True
```


## See also


#### Concepts


[Field Object](field-object-word.md)

