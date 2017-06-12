---
title: Field.DoClick Method (Word)
keywords: vbawd10.chm154075240
f1_keywords:
- vbawd10.chm154075240
ms.prod: word
api_name:
- Word.Field.DoClick
ms.assetid: 04b94737-0f7f-9086-07ff-555e416f2acf
ms.date: 06/08/2017
---


# Field.DoClick Method (Word)

Clicks the specified field.


## Syntax

 _expression_ . **DoClick**

 _expression_ Required. A variable that represents a **[Field](field-object-word.md)** object.


## Remarks

If the field is a GOTOBUTTON field, this method moves the insertion point to the specified location or selects the specified bookmark. If the field is a MACROBUTTON field, this method runs the specified macro. If the field is a HYPERLINK field, this method jumps to the target location.


## Example

If the first field in the selection is a GOTOBUTTON field, this example clicks it (the insertion point is moved to the specified location, or the specified bookmark is selected).


```vb
Dim fldTemp 
 
Set fldTemp = Selection.Fields(1) 
If fldTemp.Type = wdFieldGoToButton Then fldTemp.DoClick
```


## See also


#### Concepts


[Field Object](field-object-word.md)

