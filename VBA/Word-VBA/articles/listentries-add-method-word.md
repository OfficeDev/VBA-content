---
title: ListEntries.Add Method (Word)
keywords: vbawd10.chm153354341
f1_keywords:
- vbawd10.chm153354341
ms.prod: word
api_name:
- Word.ListEntries.Add
ms.assetid: 02e51c84-a95e-3058-e1b5-7258ac7bc65b
ms.date: 06/08/2017
---


# ListEntries.Add Method (Word)

Returns a  **ListEntry** object that represents an item added to a drop-down form field.


## Syntax

 _expression_ . **Add**( **_Name_** , **_Index_** )

 _expression_ Required. A variable that represents a **[ListEntries](listentries-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the drop-down form field item.|
| _Index_|Optional| **Variant**|A number that represents the position of the item in the list.|

### Return Value

ListEntry


## Example

This example inserts a drop-down form field in the active document and then adds the items Red, Blue, and Green to the form field.


```vb
Set myField = ActiveDocument.FormFields.Add(Range:= _ 
 Selection.Range, Type:= wdFieldFormDropDown) 
With myField.DropDown.ListEntries 
 .Add Name:="Red" 
 .Add Name:="Blue" 
 .Add Name:="Green" 
End With
```


## See also


#### Concepts


[ListEntries Collection Object](listentries-object-word.md)

