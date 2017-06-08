---
title: ListBox.Text Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 8001cbd2-b00c-7a91-9ee6-d367ff94868b
ms.date: 06/08/2017
---


# ListBox.Text Property (Outlook Forms Script)

Returns or sets a  **String** that specifies text in a **[ListBox](listbox-object-outlook-forms-script.md)**, changing the selected row in the control. Read/write.


## Syntax

 _expression_. **Text**

 _expression_A variable that represents a  **ListBox** object.


## Remarks

The default value is a zero-length string ("").

The value of  **Text** must match an existing list entry. Specifying a value that does not match an existing list entry causes an error.

You cannot use  **Text** to change the value of an entry in a **ListBox**; use the  **[Column](listbox-column-property-outlook-forms-script.md)** or **[List](listbox-list-property-outlook-forms-script.md)** property for this purpose.

The  **[ForeColor](listbox-forecolor-property-outlook-forms-script.md)** property determines the color of the text.


