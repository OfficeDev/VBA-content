---
title: ContentControlListEntry.Delete Method (Word)
keywords: vbawd10.chm147456106
f1_keywords:
- vbawd10.chm147456106
ms.prod: word
api_name:
- Word.ContentControlListEntry.Delete
ms.assetid: fa28888a-6542-9216-e444-d43b2464cf65
ms.date: 06/08/2017
---


# ContentControlListEntry.Delete Method (Word)

Deletes the specified item in a combo box or drop-down list content control.


## Syntax

 _expression_ . **Delete**

 _expression_ An expression that returns a **ContentControlListEntry** object.


## Example

The following example deletes a drop-down list item if the display text of the item is "Other".


```vb
Dim objCC As ContentControl 
Dim objCL As ContentControlListEntry 
 
For Each objCC In ActiveDocument.ContentControls 
 If objCC.Type = wdContentControlComboBox Or _ 
 objCC.Type = wdContentControlDropdownList Then 
 For Each objCL In objCC.DropdownListEntries 
 If objCL.Text = "Other" Then objCL.Delete 
 Next 
 End If 
Next 
 
```


## See also


#### Concepts


[ContentControlListEntry Object](contentcontrollistentry-object-word.md)

