---
title: WebListBoxItems.AddItem Method (Publisher)
keywords: vbapb10.chm4128772
f1_keywords:
- vbapb10.chm4128772
ms.prod: publisher
api_name:
- Publisher.WebListBoxItems.AddItem
ms.assetid: 1c3af4d1-ed0b-60c6-b607-17712612cec2
ms.date: 06/08/2017
---


# WebListBoxItems.AddItem Method (Publisher)

Adds list items to a Web list box control.


## Syntax

 _expression_. **AddItem**( **_Item_**,  **_Index_**,  **_SelectState_**,  **_ItemValue_**)

 _expression_A variable that represents a  **WebListBoxItems** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Item|Required| **String**|The name of the item as it appears in the list.|
|Index|Optional| **Long**|The number of the list item. If Index is not specified or it is out of range of the indices of existing list box items, the new item will be added to the end of the list box. Otherwise the new item will be inserted at the position specified by Index and the index position of all items after it will be increased by one.|
|SelectState|Optional| **Boolean**| **True** if the item is selected when the list box is initially displayed. Default value is **False**.|
|ItemValue|Optional| **String**|The value of the list box item. If not specified, the new item's value will be the same as the item name.|

## Remarks

When you programmatically create a new Web list box, it contains three items. Use the  **[Delete](weblistboxitems-delete-method-publisher.md)** method to remove them from the list.


## Example

This example creates a new list box control in the active publication, removes the three default list items, and then adds several items to it.


```vb
Sub AddListBoxItems() 
 Dim intCount As Integer 
 With ActiveDocument.Pages(1).Shapes.AddWebControl _ 
 (Type:=pbWebControlListBox, Left:=100, _ 
 Top:=100, Width:=150, Height:=100) 
 With .WebListBox.ListBoxItems 
 For intCount = 1 To .Count 
 .Delete (1) 
 Next 
 .AddItem Item:="Green" 
 .AddItem Item:="Yellow" 
 .AddItem Item:="Red" 
 .AddItem Item:="Blue" 
 .AddItem Item:="Purple" 
 .AddItem Item:="Chartreuse" 
 .AddItem Item:="Pink" 
 .AddItem Item:="Olive" 
 End With 
 End With 
End Sub
```


