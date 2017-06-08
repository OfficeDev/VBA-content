---
title: Range.FormFields Property (Word)
keywords: vbawd10.chm157155393
f1_keywords:
- vbawd10.chm157155393
ms.prod: word
api_name:
- Word.Range.FormFields
ms.assetid: 9777dc22-1fe5-c442-a4bf-e3dae4549168
ms.date: 06/08/2017
---


# Range.FormFields Property (Word)

Returns a  **[FormFields](formfields-object-word.md)** collection that represents all the form fields in the range. Read-only.


## Syntax

 _expression_ . **FormFields**

 _expression_ A variable that represents a **[Range](range-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example retrieves the type of the first form field in section two.


```vb
myType = ActiveDocument.Sections(2).Range.FormFields(1).Type 
Select Case myType 
 Case wdFieldFormTextInput 
 thetype = "TextBox" 
 Case wdFieldFormDropDown 
 thetype = "DropDown" 
 Case wdFieldFormCheckBox 
 thetype = "CheckBox" 
End Select
```


## See also


#### Concepts


[Range Object](range-object-word.md)

