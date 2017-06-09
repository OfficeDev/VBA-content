---
title: TextBox.IntegralHeight Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: d1ba0257-0c9e-6830-ee81-d8849c9b051a
ms.date: 06/08/2017
---


# TextBox.IntegralHeight Property (Outlook Forms Script)

Returns or sets a  **Boolean** that specifies whether a **[TextBox](textbox-object-outlook-forms-script.md)** displays full lines of text or partial lines. Read/write.


## Syntax

 _expression_. **IntegralHeight**

 _expression_A variable that represents a  **TextBox** object.


## Remarks

 **True** indicates that the text box resizes itself to display only complete items (default). **False** indicates that the text box does not resize itself even if the item is too tall to display completely.

The  **IntegralHeight** property relates to the height of the text box, just as the **[AutoSize](textbox-autosize-property-outlook-forms-script.md)** property relates to the width of the text box.

If  **IntegralHeight** is **True**, the text box automatically resizes when necessary to show full rows. If  **False**, the text box remains a fixed size; if items are taller than the available space in the text box, the entire item is not shown.


