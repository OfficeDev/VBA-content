---
title: ContentControls.Add Method (Word)
keywords: vbawd10.chm157745153
f1_keywords:
- vbawd10.chm157745153
ms.prod: word
api_name:
- Word.ContentControls.Add
ms.assetid: a9b612a6-6dcb-a74a-0b87-c112f51e2dcc
ms.date: 06/08/2017
---


# ContentControls.Add Method (Word)

Adds a new content control, of the type specified, into the active document and returns a  **[ContentControl](contentcontrol-object-word.md)** object that represents the new content control.


## Syntax

 _expression_ . **Add**( **_Type_** , **_Range_** )

 _expression_ An expression that returns a **ContentControls** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Optional| **[WdContentControlType](wdcontentcontroltype-enumeration-word.md)**|Specifies the type of content control to insert into the active document. If omitted, Microsoft Word inserts a rich-text content control.|
| _Range_|Optional| **Variant**|Specifies where in the active document to place the content control. If omitted, Word places the content control at the position of the insertion point or replaces the current selection.|

### Return Value

ContentControl


## Remarks

You can nest content controls only within rich-text content controls, building block gallery content controls, and group content controls. If the insertion point or current selection is inside a content control of a different type, this method raises an error. In this case, you can either move the insertion point or use the Range parameter to specify a location within the document.


## Example

The following example creates a new drop-down list content control and adds several items to the list.


```vb
Dim objCC As ContentControl 
 
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
objCC.DropdownListEntries.Add "Other"
```


## See also


#### Concepts


[ContentControls Collection](contentcontrols-object-word.md)

