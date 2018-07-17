---
title: ComboBox.Value Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: a81934d0-50b5-aa2d-f45b-ea8b826bcea9
ms.date: 06/08/2017
---


# ComboBox.Value Property (Outlook Forms Script)

Returns or sets a  **Variant** that specifies the value in the **[BoundColumn](combobox-boundcolumn-property-outlook-forms-script.md)** of the currently selected rows. Read/write.


## Syntax

 _expression_. **Value**

 _expression_A variable that represents a  **ComboBox** object.


## Remarks

Changing the contents of  **Value** does not change the value of **BoundColumn**. To add or delete entries in a  **ComboBox**, you can use the  **[AddItem](combobox-additem-method-outlook-forms-script.md)** or **[RemoveItem](combobox-removeitem-method-outlook-forms-script.md)** method.


