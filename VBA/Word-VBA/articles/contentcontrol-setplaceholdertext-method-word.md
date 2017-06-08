---
title: ContentControl.SetPlaceholderText Method (Word)
keywords: vbawd10.chm266534923
f1_keywords:
- vbawd10.chm266534923
ms.prod: word
api_name:
- Word.ContentControl.SetPlaceholderText
ms.assetid: d2684e44-61f0-e0bf-36bc-6a5eabed1b82
ms.date: 06/08/2017
---


# ContentControl.SetPlaceholderText Method (Word)

Sets the placeholder text that displays in the content control until a user enters their own text.


## Syntax

 _expression_ . **SetPlaceholderText**( **_BuildingBlock_** , **_Range_** , **_Text_** )

 _expression_ An expression that returns a **ContentControl** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _BuildingBlock_|Optional| **BuildingBlock**|Specifies a  **BuildingBlock** object that contains the contents of the placeholder text.|
| _Range_|Optional| **Range**|Specifies a  **Range** object that contains the contents of the placeholder text.|
| _Text_|Optional| **String**|Specifies the contents of the placeholder text.|

## Remarks

Only one of the parameters is used when specifying placeholder text. If more than one parameter is used, Microsoft Word uses the text specified in the first parameter. If all parameters are omitted, the placeholder text is blank.


## Example

The following example inserts a new drop-down list content control into the active document, sets the title and placeholder text, and then inserts several new items into the list.


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
objCC.DropdownListEntries.Add "Other"
```


## See also


#### Concepts


[ContentControl Object](contentcontrol-object-word.md)

