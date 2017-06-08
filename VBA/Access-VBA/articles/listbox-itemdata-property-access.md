---
title: ListBox.ItemData Property (Access)
keywords: vbaac10.chm11209
f1_keywords:
- vbaac10.chm11209
ms.prod: access
api_name:
- Access.ListBox.ItemData
ms.assetid: a0c1ab24-089e-6279-96dc-ef490096d06a
ms.date: 06/08/2017
---


# ListBox.ItemData Property (Access)

The  **ItemData** property returns the data in the bound column for the specified row in a list box. Read-only **Variant**.


## Syntax

 _expression_. **ItemData**( ** _Index_** )

 _expression_ A variable that represents a **ListBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Long**|The row in the combo box or list box containing the data you want to return. Rows in combo and list boxes are indexed starting with zero. For example, to return the item in the sixth row of a combo box, you'd specify 5 for the  _rowindex_ argument.|

## Remarks

The  **ItemData** property enables you to iterate through the list of entries in a combo box or list box. For example, suppose you wanted to iterate through all of the items in a list box to search for a particular entry. You can use the ListCount property to determine the number of rows in the list box, and then use the **ItemData** property to return the data for the bound column in each row.

You can also use the  **ItemData** property to return data only from selected rows in a list box. You can iterate through the ItemsSelected collection to determine which row or rows in the list box have been selected, and use the **ItemData** property to return the data in those rows. You must set the **MultiSelect** property of the list box to Simple or Extended to enable the user to select more than one row at a time.

You can use the Column property to return data from a specified row and column, even if the specified column isn't the bound column.


## Example

The following example prints the value of the bound column for each selected row in a list box EmployeeList on an Employees form. The list box's  **MultiSelect** property must be set to Simple or Extended.


```vb
Sub RowsSelected() 
 Dim ctlList As Control, varItem As Variant 
 
 ' Return Control object variable pointing to list box. 
 Set ctlList = Forms!Employees!EmployeeList 
 ' Enumerate through selected items. 
 For Each varItem in ctlList.ItemsSelected 
 ' Print value of bound column. 
 Debug.Print ctlList.ItemData(varItem) 
 Next varItem 
End Sub
```


## See also


#### Concepts


[ListBox Object](listbox-object-access.md)

