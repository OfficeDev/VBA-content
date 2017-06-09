---
title: ComboBox.SelStart Property (Outlook Forms Script)
keywords: olfm10.chm2001880
f1_keywords:
- olfm10.chm2001880
ms.prod: outlook
ms.assetid: cf739c9f-6c3a-d4fd-780b-6e6ee4559ec9
ms.date: 06/08/2017
---


# ComboBox.SelStart Property (Outlook Forms Script)

Returns or sets a  **Long** that represents the starting point of selected text, or the insertion point if no text is selected. Read/write.


## Syntax

 _expression_. **SelStart**

 _expression_A variable that represents a  **ComboBox** object.


## Remarks

For  **[SelLength](combobox-sellength-property-outlook-forms-script.md)** and **SelStart**, the valid range of settings is 0 to the total number of characters in the edit area of a  **[ComboBox](combobox-object-outlook-forms-script.md)**. The default value is zero.

The  **SelStart** property is always valid, even when the control does not have focus. Setting **SelStart** to a value less than zero creates an error. Attempting to set **SelStart** to a value greater than the number of characters available in a control results in a value equal to the number of characters in the control.

Changing the value of  **SelStart** cancels any existing selection in the control, places an insertion point in the text, and sets the **SelLength** property to zero.


