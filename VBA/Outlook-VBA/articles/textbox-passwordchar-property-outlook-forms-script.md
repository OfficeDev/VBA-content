---
title: TextBox.PasswordChar Property (Outlook Forms Script)
keywords: olfm10.chm2001690
f1_keywords:
- olfm10.chm2001690
ms.prod: outlook
ms.assetid: f9f80fb8-3c93-86fa-c717-e3bf4bde29fd
ms.date: 06/08/2017
---


# TextBox.PasswordChar Property (Outlook Forms Script)

Returns or sets a  **String** that specifies a placeholder character to be displayed instead of the characters actually entered in a **[TextBox](textbox-object-outlook-forms-script.md)**. Read/write.


## Syntax

 _expression_. **PasswordChar**

 _expression_A variable that represents a  **TextBox** object.


## Remarks

You can use the  **PasswordChar** property to protect sensitive information, such as passwords or security codes. The value of **PasswordChar** is the character (usually an asterisk) that appears in a control instead of the actual characters that the user types. If you don't specify a character, the control displays the characters that the user types.


