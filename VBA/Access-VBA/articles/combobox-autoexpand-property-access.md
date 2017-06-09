---
title: ComboBox.AutoExpand Property (Access)
keywords: vbaac10.chm11388,vbaac10.chm4275
f1_keywords:
- vbaac10.chm11388,vbaac10.chm4275
ms.prod: access
api_name:
- Access.ComboBox.AutoExpand
ms.assetid: 0b3fabf8-4004-0868-3ddc-aef297514324
ms.date: 06/08/2017
---


# ComboBox.AutoExpand Property (Access)

You can use the  **AutoExpand** property to specify whether Microsoft Access automatically fills the text box portion of a combo box with a value from the combo box list that matches the characters you enter as you type in the combo box. This lets you quickly enter an existing value in a combo box without displaying the list box portion of the combo box. Read/write **Boolean**.


## Syntax

 _expression_. **AutoExpand**

 _expression_ A variable that represents a **ComboBox** object.


## Remarks

When you enter characters in the text box portion of a combo box, Office Access 2007 searches the values in the list to find those that match the characters you have typed. If the  **AutoExpand** property is set to Yes, Microsoft Access automatically displays the first underlying value that matches the characters entered so far.

When the  **[LimitToList](combobox-limittolist-property-access.md)** property is set to **Yes** and the combo box list is dropped down, Microsoft Access selects matching values in the list as the user enters characters in the text box portion of the combo box, even if the **AutoExpand** property is set to **No**. If the user presses ENTER or moves to another control or record, the selected value appears in the combo box.


## See also


#### Concepts


[ComboBox Object](combobox-object-access.md)

