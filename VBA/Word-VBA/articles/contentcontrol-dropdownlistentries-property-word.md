---
title: ContentControl.DropdownListEntries Property (Word)
keywords: vbawd10.chm266534921
f1_keywords:
- vbawd10.chm266534921
ms.prod: word
api_name:
- Word.ContentControl.DropdownListEntries
ms.assetid: 4434c4cc-53f4-758d-5a9e-3a9aa2f74a05
ms.date: 06/08/2017
---


# ContentControl.DropdownListEntries Property (Word)

Returns a  **[ContentControlListEntries](contentcontrollistentries-object-word.md)** collection that represents the items in a drop-down list content control or in a combo box content control. Read-only.

For Office 2016, returns dropdown entries for accessing individual list items within a collection, with the exception of SharePoint lookups.

## Syntax

 _expression_ . **DropdownListEntries**

 _expression_ An expression that returns a **ContentControl** object.


## Example

The following example inserts a new drop-down list content control into the active document, sets the title and placeholder text, and then adds several new items to the list.


```vb
Dim objCC As ContentControl 
Dim objMap As XMLMapping 
 
Set objCC = ActiveDocument.ContentControls.Add(wdContentControlDropdownList) 
objCC.Title = "My Favorite Animal" 
objCC.SetPlaceholderText , , "Select your favorite animal " 
 
'List entries 
objCC.DropdownListEntries.Add "Cat" 
objCC.DropdownListEntries.Add "Dog" 
objCC.DropdownListEntries.Add "Horse" 
objCC.DropdownListEntries.Add "Monkey" 
objCC.DropdownListEntries.Add "Snake" 
objCC.DropdownListEntries.Add("Other")
```


## See also


#### Concepts


[ContentControl Object](contentcontrol-object-word.md)

