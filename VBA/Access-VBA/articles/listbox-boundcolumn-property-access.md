---
title: ListBox.BoundColumn Property (Access)
keywords: vbaac10.chm11227
f1_keywords:
- vbaac10.chm11227
ms.prod: access
api_name:
- Access.ListBox.BoundColumn
ms.assetid: f6a742a4-40ff-bb83-8946-7e8bb71e5690
ms.date: 06/08/2017
---


# ListBox.BoundColumn Property (Access)

When you make a selection from a list box, the  **BoundColumn** property tells Microsoft Access which column's values to use as the value of the control. If the control is bound to a field, the value in the column specified by the **BoundColumn** property is stored in the field named in the **ControlSource** property. Read/write **Long**.


## Syntax

 _expression_. **BoundColumn**

 _expression_ A variable that represents a **ListBox** object.


## Remarks

The  **BoundColumn** property uses the following settings.



|**Setting**|**Description**|
|:-----|:-----|
|0|The  **ListIndex** property value, rather than the column value, is stored in the current record. The **ListIndex** property value of the first row is 0, the second row is 1, and so on. Microsoft Access sets the **ListIndex** property when an item is selected from a list box or the list box portion of a combo box. Setting the **BoundColumn** property to 0 and using the **ListIndex** property value of the control might be useful if, for example, you are only interested in storing a sequence of numbers.|
|1 or greater|(Default is 1) The value in the specified column becomes the control's value. If the control is bound to a field, then this setting is stored in that field in the current record. The BoundColumn property can't be set to a value larger than the setting of the ColumnCount property.|
For table fields , you can set this property on the  **Lookup** tab in the Field Properties section of table Design view for fields with the **DisplayControl** property set to Combo Box or List Box.

In Visual Basic, set the  **BoundColumn** property by using a number or a numeric expression equal to a value from 0 to the setting of the **ColumnCount** property.


 **Note**  If the bound column is not the same as the leftmost visible column in the control (or if you set the  **BoundColumn** property to 0), the **LimitToList** property is set to Yes.

Microsoft Access uses zero-based numbers to refer to columns in the  **Column** property. That is, the first column is referenced by using the expression `Column(0)`; the second column is referenced by using the expression  `Column(1)`; and so on. However, the  **BoundColumn** property uses 1-based numbers to refer to the columns. This means that if the **BoundColumn** property is set to 1, you could access the value stored in that column by using the expression `Column(0)`.

If the  **AutoExpand** property is set to Yes, Microsoft Access automatically fills in a value in the text box portion of the combo box that matches a value in the combo box list as you type.


## See also


#### Concepts


[ListBox Object](listbox-object-access.md)

