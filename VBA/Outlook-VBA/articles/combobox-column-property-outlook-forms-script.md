---
title: ComboBox.Column Property (Outlook Forms Script)
keywords: olfm10.chm2000930
f1_keywords:
- olfm10.chm2000930
ms.prod: outlook
ms.assetid: f00c388f-fe1f-5458-281f-4bfa549291d5
ms.date: 06/08/2017
---


# ComboBox.Column Property (Outlook Forms Script)

Returns or sets a  **Variant** that represents a single value, a column of values, or a two-dimensional array to load into a **[ComboBox](combobox-object-outlook-forms-script.md)**. Read/write.


## Syntax

 _expression_. **Column**( **_pvargColumn_**,  **_pvargIndex_**)

 _expression_A variable that represents a  **ComboBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|pvargColumn|Optional| **Variant**|An integer with a range from 0 to one less than the total number of columns.|
|pvargIndex|Optional| **Variant**|An integer with a range from 0 to one less than the total number of rows.|

## Remarks

If you specify both the column and row values,  **Column** reads or writes a specific item.

If you specify only the column value, the  **Column** property reads or writes the specified column in the current row of the object. For example, `MyComboBox.Column (3)` reads or writes the third column in MyComboBox.

 **Column** returns a **Variant** from the cursor. When a built-in cursor provides the value for **Variant** (such as when using the **[AddItem](combobox-additem-method-outlook-forms-script.md)** method), the value is a **String**. When an external cursor provides the value for  **Variant**, formatting associated with the data is not included in the  **Variant**.

You can use  **Column** to assign the contents of a combo box to another control, such as a text box.

If the user makes no selection when you refer to a column in a combo box or list box, the  **Column** setting is **Null**. You can check for this condition by using the  **IsNull** function.

You can also use  **Column** to copy an entire two-dimensional array of values to a control. This syntax lets you quickly load a list of choices rather than individually loading each element of the list using **AddItem**.

When copying data from a two-dimensional array,  **Column** transposes the contents of the array in the control so that the contents of `ComboBox1.Column(X, Y)` is the same as `MyArray(Y, X)`. You can also use the  **[List](combobox-list-property-outlook-forms-script.md)** property to copy an array without transposing it.


