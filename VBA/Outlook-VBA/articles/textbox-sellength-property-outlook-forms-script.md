---
title: TextBox.SelLength Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 97d11d04-a1d9-4251-01fc-a64f6d1293ee
ms.date: 06/08/2017
---


# TextBox.SelLength Property (Outlook Forms Script)

Returns or sets a  **Long** that represents the number of characters selected in a **[TextBox](textbox-object-outlook-forms-script.md)**. Read/write.


## Syntax

 _expression_. **SelLength**

 _expression_A variable that represents a  **TextBox** object.


## Remarks

For  **SelLength** and **[SelStart](textbox-selstart-property-outlook-forms-script.md)**, the valid range of settings is 0 to the total number of characters in the edit area of a  **TextBox**.

The  **SelLength** property is always valid, even when the control does not have focus. Setting **SelLength** to a value less than zero creates an error. Attempting to set **SelLength** to a value greater than the number of characters available in a control results in a value equal to the number of characters in the control.

Changing the value of the  **SelStart** property cancels any existing selection in the control, places an insertion point in the text, and sets **SelLength** to zero.

The default value, zero, means that no text is currently selected.


