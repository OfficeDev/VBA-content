---
title: ListBox.ListIndex Property (Access)
keywords: vbaac10.chm11275
f1_keywords:
- vbaac10.chm11275
ms.prod: access
api_name:
- Access.ListBox.ListIndex
ms.assetid: 4d75f753-f6a9-d362-7281-2c0460f4272b
ms.date: 06/08/2017
---


# ListBox.ListIndex Property (Access)

You can use the  **ListIndex** property to determine which item is selected in a list box. Read/write **Long**.


## Syntax

 _expression_. **ListIndex**

 _expression_ A variable that represents a **ListBox** object.


## Remarks

The  **ListIndex** property is an integer from 0 to the total number of items in a list box or combo box minus 1. Microsoft Access sets the **ListIndex** property value when an item is selected in a list box or list box portion of a combo box. The **ListIndex** property value of the first item in a list is 0, the value of the second item is 1, and so on.

This property is available only by using a macro or Visual Basic . You can read this property only in Form view and Datasheet view. This property is read-only and isn't available in other views.

The  **ListIndex** property value is also available by setting the **BoundColumn** property to 0 for a combo box or list box. If the **BoundColumn** property is set to 0, the underlying table field to which the combo box or list box is bound will contain the same value as the **ListIndex** property setting.

List boxes also have a  **MultiSelect** property that allows the user to select multiple items from the control. When multiple selections are made in a list box, you can determine which items are selected by using the **Selected** property of the control. The **Selected** property is an array of values from 0 to the **ListCount** property value minus 1. For each item in the list box the **Selected** property will be **True** if the item is selected and **False** if it is not selected.

The  **ItemsSelected** collection also provides a way to access data in the selected rows of a list box or combo box.


## See also


#### Concepts


[ListBox Object](listbox-object-access.md)

