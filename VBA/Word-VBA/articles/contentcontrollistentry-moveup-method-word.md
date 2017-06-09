---
title: ContentControlListEntry.MoveUp Method (Word)
keywords: vbawd10.chm147456107
f1_keywords:
- vbawd10.chm147456107
ms.prod: word
api_name:
- Word.ContentControlListEntry.MoveUp
ms.assetid: e67c7c3c-fdf0-64b4-7e93-7e6f7a47c9bd
ms.date: 06/08/2017
---


# ContentControlListEntry.MoveUp Method (Word)

Moves an item in a drop-down list or combo box content control up one item, so that it is before the item that originally preceded it.


## Syntax

 _expression_ . **MoveUp**

 _expression_ An expression that returns a **ContentControlListEntry** object.


## Example

The following example moves the last item in the drop-down list or combo box content control up, so that it becomes the first item.


```vb
Dim objCC As ContentControl 
Dim objCL As ContentControlListEntry 
Dim intCount As Integer 
 
Set objCC = ActiveDocument.ContentControls.Item(3) 
 
If objCC.Type = wdContentControlComboBox Or _ 
 objCC.Type = wdContentControlDropdownList Then 
 
 Set objCL = objCC.DropdownListEntries.Item(objCC.DropdownListEntries.Count) 
 
 For intCount = 1 To objCC.DropdownListEntries.Count 
 objCL.MoveUp 
 Next 
 
End If
```


## See also


#### Concepts


[ContentControlListEntry Object](contentcontrollistentry-object-word.md)

