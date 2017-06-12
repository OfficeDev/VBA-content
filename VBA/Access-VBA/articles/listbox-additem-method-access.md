---
title: ListBox.AddItem Method (Access)
keywords: vbaac10.chm11301
f1_keywords:
- vbaac10.chm11301
ms.prod: access
api_name:
- Access.ListBox.AddItem
ms.assetid: dab0c3e4-8ecc-774b-4c7e-f973eb4c1516
ms.date: 06/08/2017
---


# ListBox.AddItem Method (Access)

Adds a new item to the list of values displayed by the specified list box control.


## Syntax

 _expression_. **AddItem**( ** _Item_**, ** _Index_** )

 _expression_ A variable that represents a **ListBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Item_|Required|**String**|The display text for the new item.|
| _Index_|Optional|**Variant**|The position of the item in the list. If this argument is omitted, the item is added to the end of the list.|

## Remarks

The  **RowSourceType** property of the specified control must be set to "Value List".

This method is only valid for list box or combo box controls on forms.

List item numbers start from zero. If the value of the  _Item_ argument doesn't correspond to an existing item number, an error occurs.

For multiple-column lists, use semicolons to delimit the strings for each column (for example, "1010;red;large" for a three-column list). If the  _Item_ argument contains fewer strings than columns in the control, items will be added starting with the left-most column. If the _Item_ argument contains more strings than columns in the control, the extra strings are ignored.

Use the  **RemoveItem** method to remove items from the list of values.


## Example

This example adds an item to the end of the list in a list box control. For the function to work, you must pass it a  **ListBox** object representing a list box control on a form and a **String** value representing the text of the item to be added.


```vb
Function AddItemToEnd(ctrlListBox As ListBox, _ 
 ByVal strItem As String) 
 
 ctrlListBox.AddItem Item:=strItem 
 
End Function
```

This example adds an item to the beginning of the list in a combo box control. For the function to work, you must pass it a  **ComboBox** object representing a combo box control on a form and a **String** value representing the text of the item to be added.




```vb
Function AddItemToBeginning(ctrlComboBox As ComboBox, _ 
 ByVal strItem As String) 
 
 ctrlComboBox.AddItem Item:=strItem, Index:=0 
 
End Function
```


## See also


#### Concepts


[ListBox Object](listbox-object-access.md)

