---
title: Field.ShowCodes Property (Word)
keywords: vbawd10.chm154075145
f1_keywords:
- vbawd10.chm154075145
ms.prod: word
api_name:
- Word.Field.ShowCodes
ms.assetid: 36871ffb-b307-c36e-5896-74fba6feb524
ms.date: 06/08/2017
---


# Field.ShowCodes Property (Word)

 **True** if field codes are displayed for the specified field instead of field results. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowCodes**

 _expression_ An expression that returns a **[Field](field-object-word.md)** object.


## Example

This example selects the next field and displays the field codes.


```vb
With Selection 
 .GoTo What:=wdGoToField 
 .Expand Unit:=wdWord 
 If .Fields.Count = 1 Then .Fields(1).ShowCodes = True 
End With
```

This example updates and displays the result of the first field in the active document.




```vb
If ActiveDocument.Fields.Count >= 1 Then 
 With ActiveDocument.Fields(1) 
 .Update 
 .ShowCodes = False 
 End With 
End If
```


## See also


#### Concepts


[Field Object](field-object-word.md)

