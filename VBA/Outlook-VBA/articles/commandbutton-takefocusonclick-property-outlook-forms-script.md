---
title: CommandButton.TakeFocusOnClick Property (Outlook Forms Script)
keywords: olfm10.chm2001340
f1_keywords:
- olfm10.chm2001340
ms.prod: outlook
ms.assetid: b8842b50-4be8-c366-8978-8a6c97907e33
ms.date: 06/08/2017
---


# CommandButton.TakeFocusOnClick Property (Outlook Forms Script)

Returns or sets a  **Boolean** that specifies whether a control takes the focus when clicked. Read/write.


## Syntax

 _expression_. **TakeFocusOnClick**

 _expression_A variable that represents a  **CommandButton** object.


## Remarks

 **True** if the button takes the focus when clicked (default). **False** if the button does not take the focus when clicked.

The  **TakeFocusOnClick** property defines only what happens when the user clicks a control. If the user tabs to the control, the control takes the focus regardless of the value of **TakeFocusOnClick**.

Use this property to complete actions that affect a control without requiring that control to give up focus. For example, assume your form includes a  **[TextBox](textbox-object-outlook-forms-script.md)** and a **[CommandButton](commandbutton-object-outlook-forms-script.md)** that checks for correct spelling of text. You would like to be able to select text in the **TextBox**, then click the  **CommandButton** and run the spelling checker without taking focus away from the **TextBox**. You can do this by setting the  **TakeFocusOnClick** property of the **CommandButton** to **False**.


