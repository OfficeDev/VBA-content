---
title: Explorer.IsItemSelectableInView Method (Outlook)
keywords: vbaol11.chm3308
f1_keywords:
- vbaol11.chm3308
ms.prod: outlook
api_name:
- Outlook.Explorer.IsItemSelectableInView
ms.assetid: a2ec8bbb-0f24-6db6-05a8-1b8375b71da7
ms.date: 06/08/2017
---


# Explorer.IsItemSelectableInView Method (Outlook)

Returns a value that indicates whether the specified Microsoft Outlook item can be selected in the current view of the active explorer.


## Syntax

 _expression_ . **IsItemSelectableInView**( **_Item_** )

 _expression_ A variable that represents an **[Explorer](explorer-object-outlook.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Item_|Required| **Object**|The item that is being considered for selection.|

### Return Value

A  **Boolean** value that indicates whether the specified item can be selected in the current view.


## Remarks

Returns  **True** if the item can be selected in the current view; otherwise returns **False** .

 The method returns **True** or **False** depending on whether the item is selectable in the view. It does not indicate whether the item is visible in the view. If the item is contained within a collapsed group in the view, the method returns **False** .

If in-cell editing is turned on for the current view for the active explorer and in-cell editing is in progress on the item, the method returns  **False** . If the current view is a table view, you can check for in-cell editing by using the **[AllowInCellEditing](tableview-allowincellediting-property-outlook.md)** property of the **[TableView](tableview-object-outlook.md)** object. Similarly, if the current view is a card view, you can use the **[AllowInCellEditing](cardview-allowincellediting-property-outlook.md)** property of the **[CardView](cardview-object-outlook.md)** object.

When you specify an item in a recurring appointment or task as argument to the  **IsItemSelectableInView** method, make sure that before you pass the argument, you obtain an instance of the occurrence by first expanding the recurrences, using the **[IncludeRecurrences](items-includerecurrences-property-outlook.md)** property and the **[Items](items-object-outlook.md)** collection. If you do not expand the recurrences and obtain an occurrence in the series, you would be passing an instance variable that represents the appointment or task series, and the **IsItemSelectableInView** method would be operating on the series instead of the occurrence.

The  **IsItemSelectableInView** method raises an error if the current view is a conversation view.


## See also


#### Concepts


[Explorer Object](explorer-object-outlook.md)

