---
title: PasswordChar Property
keywords: fm20.chm5225076
f1_keywords:
- fm20.chm5225076
ms.prod: office
api_name:
- Office.PasswordChar
ms.assetid: 2dd645b2-fe8d-a644-b796-e0595627cbb8
ms.date: 06/08/2017
---


# PasswordChar Property



Specifies whether [placeholder](glossary-vba.md) characters are displayed instead of the characters actually entered in a **TextBox**.
 **Syntax**
 _object_. **PasswordChar** [= _String_ ]
The  **PasswordChar** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _String_|Optional. A string expression specifying the placeholder character.|
 **Remarks**
You can use the  **PasswordChar** property to protect sensitive information, such as passwords or security codes. The value of **PasswordChar** is the character that appears in a control instead of the actual characters that the user types. If you don't specify a character, the control displays the characters that the user types.

