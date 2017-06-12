---
title: TextBox.AutoWordSelect Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 00fb7b7b-e7ab-a996-765d-04207d6ba995
ms.date: 06/08/2017
---


# TextBox.AutoWordSelect Property (Outlook Forms Script)

Returns or sets a  **Boolean** that specifies whether the basic unit used to extend a selection is a word or a single character. Read/write.


## Syntax

 _expression_. **AutoWordSelect**

 _expression_A variable that represents a  **TextBox** object.


## Remarks

 **True** if uses a word as the basic unit (default), **False** if uses a character as the basic unit.

The  **AutoWordSelect** property specifies how the selection extends or contracts in the edit region of a **[TextBox](textbox-object-outlook-forms-script.md)**.

If the user places the insertion point in the middle of a word and then extends the selection while  **AutoWordSelect** is **True**, the selection includes the entire word.


