---
title: TextBox.CurX Property (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 1e367959-9f87-c79c-b816-aabf8cde2e23
ms.date: 06/08/2017
---


# TextBox.CurX Property (Outlook Forms Script)

Returns or sets a  **Long** that represents the current horizontal position of the insertion point in a multiline **[TextBox](textbox-object-outlook-forms-script.md)**. Read/write.


## Syntax

 _expression_. **CurX**

 _expression_A variable that represents a  **TextBox** object.


## Remarks

The  **CurX** property applies to a multiline **TextBox**. The return value is valid when the object has the focus.

You can use  **[CurTargetX](textbox-curtargetx-property-outlook-forms-script.md)** and **CurX** to position the insertion point as the user scrolls through the contents of a multiline **TextBox**. When the user moves the insertion point to another line of text by scrolling the content of the object,  **CurTargetX** specifies the preferred position for the insertion point. **CurX** is set to this value if the line of text is longer than the value of **CurTargetX**. Otherwise,  **CurX** is set to the end of the line of text.


