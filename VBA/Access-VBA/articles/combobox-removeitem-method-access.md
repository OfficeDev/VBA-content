---
title: ComboBox.RemoveItem Method (Access)
keywords: vbaac10.chm11477
f1_keywords:
- vbaac10.chm11477
ms.prod: access
api_name:
- Access.ComboBox.RemoveItem
ms.assetid: 9e70c221-e2fd-d006-1460-2b1902b0b0ea
ms.date: 06/08/2017
---


# ComboBox.RemoveItem Method (Access)

Removes an item from the list of values displayed by the specified combo box control.


## Syntax

 _expression_. **RemoveItem**( ** _Index_** )

 _expression_ A variable that represents a **ComboBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Variant**|The item to be removed from the list, expressed as either an item number or the list item text.|

## Remarks

This method is only valid for list box or combo box controls on forms. Also, the  **RowSourceType** property of the control must be set to "Value List".

List item numbers start from zero. If the value of the  _Index_ argument doesn't correspond to an existing item number or the text of an existing item, an error occurs.

Use the  **AddItem** method to add items to the list of values.


## Example

This example removes the specified item from the list in a list box control. For the function to work, you must pass it a  **ListBox** object representing a list box control on a form and a **Variant** value representing the item to be removed.


```vb
Function RemoveListItem(ctrlListBox As ListBox, _ 
 ByVal varItem As Variant) As Boolean 
 
 ' Trap for errors. 
 On Error GoTo ERROR_HANDLER 
 
 ' Remove the list box item and set the return value 
 ' to True, indicating success. 
 ctrlListBox.RemoveItem Index:=varItem 
 RemoveListItem = True 
 
 ' Reset the error trap and exit the function. 
 On Error GoTo 0 
 Exit Function 
 
' Return False if an error occurs. 
ERROR_HANDLER: 
 RemoveListItem = False 
 
End Function
```


## See also


#### Concepts


[ComboBox Object](combobox-object-access.md)

