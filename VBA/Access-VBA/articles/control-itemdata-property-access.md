---
title: Control.ItemData Property (Access)
keywords: vbaac10.chm10141
f1_keywords:
- vbaac10.chm10141
ms.prod: access
api_name:
- Access.Control.ItemData
ms.assetid: 5eb23c40-566e-33bb-9b73-0ecc701ea5e5
ms.date: 06/08/2017
---


# Control.ItemData Property (Access)

The  **ItemData** property returns the data in the bound column for the specified row in a combo box or list box. Read-only **Variant**.


## Syntax

 _expression_. **ItemData**( ** _Index_** )

 _expression_ A variable that represents a **Control** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Long**|The row in the combo box or list box containing the data you want to return. Rows in combo and list boxes are indexed starting with zero. For example, to return the item in the sixth row of a combo box, you'd specify 5 for the  _rowindex_ argument.|

## Remarks

The  **ItemData** property enables you to iterate through the list of entries in a combo box or list box. For example, suppose you wanted to iterate through all of the items in a list box to search for a particular entry. You can use the ListCount property to determine the number of rows in the list box, and then use the **ItemData** property to return the data for the bound column in each row.

You can also use the  **ItemData** property to return data only from selected rows in a list box. You can iterate through the ItemsSelected collection to determine which row or rows in the list box have been selected, and use the **ItemData** property to return the data in those rows. You must set the **MultiSelect** property of the list box to Simple or Extended to enable the user to select more than one row at a time.

You can use the Column property to return data from a specified row and column, even if the specified column isn't the bound column.


## See also


#### Concepts


[Control Object](control-object-access.md)

