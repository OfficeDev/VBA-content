---
title: AddItem Method
keywords: fm20.chm2000260
f1_keywords:
- fm20.chm2000260
ms.prod: office
api_name:
- Office.AddItem
ms.assetid: cd8ce314-7ba2-5930-5747-4eb89c649630
ms.date: 06/08/2017
---


# AddItem Method



For a single-column list box or combo box, adds an item to the list. For a multicolumn list box or combo box, adds a row to the list.
 **Syntax**
 _Variant_ = _object_. **AddItem** [ _item_ [, _varIndex_ ]]
The  **AddItem** method syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Item_|Optional. Specifies the item or row to add. The number of the first item or row is 0; the number of the second item or row is 1, and so on.|
| _varIndex_|Optional. Integer specifying the position within the object where the new item or row is placed.|
 **Remarks**
If you supply a valid value for  _varIndex_, the **AddItem** method places the item or row at that position within the list. If you omit _varIndex_, the method adds the item or row at the end of the list.
The value of  _varIndex_ must not be greater than the value of the **ListCount** property.
For a multicolumn  **ListBox** or **ComboBox**, **AddItem** inserts an entire row, that is, it inserts an item for each column of the control. To assign values to an item beyond the first column, use the **List** or **Column** property and specify the row and column of the item.
If the control is bound to data, the  **AddItem** method fails.

 **Note**  You can add more than one row at a time to a  **ComboBox** or **ListBox** by using **List**.


