---
title: ContentControlListEntry.Select Method (Word)
keywords: vbawd10.chm147456109
f1_keywords:
- vbawd10.chm147456109
ms.prod: word
api_name:
- Word.ContentControlListEntry.Select
ms.assetid: f600e267-39d9-238d-c6d2-9efba6b4044d
ms.date: 06/08/2017
---


# ContentControlListEntry.Select Method (Word)

Selects the list entry in a drop-down list or combo box content control and sets the text of the content control to the value of the item.


## Syntax

 _expression_ . **Select**

 _expression_ An expression that returns a **ContentControlListEntry** object.


## Example

The following example inserts a drop-down list content control into the active document, sets the title and placeholder text and adds several items to the list, and then selects the last item entered.


```vb
Dim objCC As ContentControl 
Dim objCE As ContentControlListEntry 
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
 
Set objCE = objCC.DropdownListEntries.Add("Other") 
objCE.Select
```


## See also


#### Concepts


[ContentControlListEntry Object](contentcontrollistentry-object-word.md)

