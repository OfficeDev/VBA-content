---
title: ContentControlListEntry.MoveDown Method (Word)
keywords: vbawd10.chm147456108
f1_keywords:
- vbawd10.chm147456108
ms.prod: word
api_name:
- Word.ContentControlListEntry.MoveDown
ms.assetid: 9b8e366e-3d04-c5d5-b9b5-0a91e10b8c1f
ms.date: 06/08/2017
---


# ContentControlListEntry.MoveDown Method (Word)

Moves an item in a drop-down list or combo box content control down one item, so that it is after the item that originally followed it.


## Syntax

 _expression_ . **MoveDown**

 _expression_ An expression that returns a **ContentControlListEntry** object.


## Example

The following example moves the first item down, so that it becomes the last item in the list of items in a drop-down list or combo box content control.


```vb
Dim objCC As ContentControl 
Dim objCL As ContentControlListEntry 
Dim intCount As Integer 
 
Set objCC = ActiveDocument.ContentControls.Item(3) 
 
If objCC.Type = wdContentControlComboBox Or _ 
 objCC.Type = wdContentControlDropdownList Then 
 
 Set objCL = objCC.DropdownListEntries.Item(1) 
 
 For intCount = 1 To objCC.DropdownListEntries.Count 
 objCL.MoveDown 
 Next 
 
End If
```


## See also


#### Concepts


[ContentControlListEntry Object](contentcontrollistentry-object-word.md)

