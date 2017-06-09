---
title: ContentControl.Title Property (Word)
keywords: vbawd10.chm266534924
f1_keywords:
- vbawd10.chm266534924
ms.prod: word
api_name:
- Word.ContentControl.Title
ms.assetid: 3bfd7bd5-2477-95ed-a334-bb3e7e450036
ms.date: 06/08/2017
---


# ContentControl.Title Property (Word)

Returns or sets a  **String** that represents the title for a content control. Read/write.


## Syntax

 _expression_ . **Title**

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

