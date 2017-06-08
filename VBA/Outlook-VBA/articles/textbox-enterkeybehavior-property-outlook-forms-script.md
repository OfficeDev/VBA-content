---
title: TextBox.EnterKeyBehavior Property (Outlook Forms Script)
keywords: olfm10.chm2001130
f1_keywords:
- olfm10.chm2001130
ms.prod: outlook
ms.assetid: 2af4a64e-4939-ae46-0d25-67fe986d413a
ms.date: 06/08/2017
---


# TextBox.EnterKeyBehavior Property (Outlook Forms Script)

Returns or sets a  **Boolean** that defines the effect of pressing **ENTER** in a **[TextBox](textbox-object-outlook-forms-script.md)**. Read/write.


## Syntax

 _expression_. **EnterKeyBehavior**

 _expression_A variable that represents a  **TextBox** object.


## Remarks

 **True** if pressing **ENTER** creates a new line, **False** if pressing **ENTER** moves the focus to the next object in the tab order (default).

The  **EnterKeyBehavior** and **[MultiLine](textbox-multiline-property-outlook-forms-script.md)** properties are closely related. The values described above only apply if **MultiLine** is **True**. If  **MultiLine** is **False**, pressing  **ENTER** always moves the focus to the next control in the tab order regardless of the value of **EnterKeyBehavior**.

The effect of pressing  **CTRL+ENTER** also depends on the value of **MultiLine**. If  **MultiLine** is **True**, pressing  **CTRL+ENTER** creates a new line regardless of the value of **EnterKeyBehavior**. If  **MultiLine** is **False**, pressing  **CTRL+ENTER** has no effect.


