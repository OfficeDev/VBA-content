---
title: WebListBoxItems.Selected Method (Publisher)
keywords: vbapb10.chm4128775
f1_keywords:
- vbapb10.chm4128775
ms.prod: publisher
api_name:
- Publisher.WebListBoxItems.Selected
ms.assetid: 2db3b8cb-2922-1cca-9613-67402772ee27
ms.date: 06/08/2017
---


# WebListBoxItems.Selected Method (Publisher)

Selects or cancels the selection of an item in a Web list box control.


## Syntax

 _expression_. **Selected**( **_Index_**,  **_SelectState_**)

 _expression_A variable that represents a  **WebListBoxItems** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Required| **Long**|The number of the Web list box item.|
|SelectState|Required| **Boolean**| **True** to select the list item.|

## Example

This example verifies that an existing Web list box control allows selecting multiple entries and then selects two items in the list.


```vb
Sub SelectListBoxItem() 
 With ActiveDocument.Pages(1).Shapes(1) _ 
 .WebListBox 
 If .MultiSelect = msoTrue Then 
 With .ListBoxItems 
 .Selected Index:=1, SelectState:=True 
 .Selected Index:=3, SelectState:=True 
 End With 
 End If 
 End With 
End Sub
```


