---
title: DropDown.ListEntries Property (Word)
keywords: vbawd10.chm153419779
f1_keywords:
- vbawd10.chm153419779
ms.prod: word
api_name:
- Word.DropDown.ListEntries
ms.assetid: 87235132-0ff6-e8d7-1efc-1df4a9816b2f
ms.date: 06/08/2017
---


# DropDown.ListEntries Property (Word)

Returns a  **[ListEntries](listentries-object-word.md)** collection that represents all the items in a **DropDown** object.


## Syntax

 _expression_ . **ListEntries**

 _expression_ An expression that returns a **[DropDown](dropdown-object-word.md)** object.


## Remarks

For information about returning a single member of a collection, see [Returning an Object from a Collection](http://msdn.microsoft.com/library/28f76384-f495-9640-a7c8-10ada3fac727%28Office.15%29.aspx).


## Example

This example retrieves the text of the active item from the drop-down form field named "DropDown1."


```vb
Set myField = ActiveDocument.FormFields("DropDown1").DropDown 
num = myField.Value 
myName = myField.ListEntries(num).Name
```

This example retrieves the total number of items in the active drop-down form field (the document should be protected for forms). If there are two or more items, this example sets the second item as the active item.




```vb
Set myField = Selection.FormFields(1) 
If myfield.Type = wdFieldFormDropDown Then 
 num = myField.DropDown.ListEntries.Count 
 If num >= 2 Then myField.DropDown.Value = 2 
End If
```


## See also


#### Concepts


[DropDown Object](dropdown-object-word.md)

