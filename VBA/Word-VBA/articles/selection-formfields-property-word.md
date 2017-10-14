---
title: Selection.FormFields Property (Word)
keywords: vbawd10.chm158662721
f1_keywords:
- vbawd10.chm158662721
ms.prod: word
api_name:
- Word.Selection.FormFields
ms.assetid: d6d5259b-9971-929f-16f7-ca2b2d585c77
ms.date: 06/08/2017
---


# Selection.FormFields Property (Word)

Returns a  **[FormFields](formfields-object-word.md)** collection that represents all the form fields in the selection. Read-only.


## Syntax

 _expression_ . **FormFields**

 _expression_ A variable that represents a **[Selection](selection-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example displays the name of the first form field in the selection.


```vb
If Selection.FormFields.Count > 0 Then 
 MsgBox Selection.FormFields(1).Name 
End If
```


## See also


#### Concepts


[Selection Object](selection-object-word.md)

