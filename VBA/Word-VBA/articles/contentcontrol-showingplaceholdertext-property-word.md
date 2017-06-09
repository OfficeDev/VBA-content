---
title: ContentControl.ShowingPlaceholderText Property (Word)
keywords: vbawd10.chm266534931
f1_keywords:
- vbawd10.chm266534931
ms.prod: word
api_name:
- Word.ContentControl.ShowingPlaceholderText
ms.assetid: 1c502641-f969-10d9-ebe5-04c85f0bfe48
ms.date: 06/08/2017
---


# ContentControl.ShowingPlaceholderText Property (Word)

Returns a  **Boolean** that indicates whether the placeholder text for the content control is displayed. Read-only.


## Syntax

 _expression_ . **ShowingPlaceholderText**

 _expression_ An expression that returns a **ContentControl** object.


## Example

The following example inserts a new drop-down list content control into the active document, sets the title and the placeholder text if the placeholder text is showing, and then adds several new items to the list.


```vb
Dim objCC As ContentControl 
Dim objMap As XMLMapping 
 
Set objCC = ActiveDocument.ContentControls.Add(wdContentControlDropdownList) 
objCC.Title = "My Favorite Animal" 
If objCC.ShowingPlaceholderText Then _ 
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

