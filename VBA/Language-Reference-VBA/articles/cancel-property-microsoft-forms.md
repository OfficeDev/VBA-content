---
title: Cancel Property (Microsoft Forms)
keywords: fm20.chm2000840
f1_keywords:
- fm20.chm2000840
ms.prod: office
ms.assetid: ac816d52-a1a3-9d64-f70a-0d96d49766a2
ms.date: 06/08/2017
---


# Cancel Property (Microsoft Forms)



Returns or sets a value indicating whether a command button is the Cancel button on a form.
 **Syntax**
 _object_. **Cancel** [= _Boolean_ ]
The  **Cancel** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Boolean_|Optional. Whether the object is the Cancel button.|
 **Settings**
The settings for  _Boolean_ are:


|**Value**|**Description**|
|:-----|:-----|
|**True**|The  **CommandButton** is the Cancel button.|
|**False**|The  **CommandButton** is not the Cancel button (default).|
 **Remarks**
A  **CommandButton** or an object that acts like a command button can be designated as the default command button. For[OLE container controls](glossary-vba.md) (Windows only), the **Cancel** property is provided only for those objects that specifically behave as command buttons.
Only one  **CommandButton** on a form can be the Cancel button. Setting **Cancel** to **True** for one command button automatically sets it to **False** for all other objects on the form. When a **CommandButton's Cancel** property is set to **True** and the form is the active form, the user can choose the command button by clicking it, pressing ESC, or pressing ENTER when the button has the[focus](vbe-glossary.md).
A typical use of  **Cancel** is to give the user the option of canceling uncommitted changes and returning the form to its previous state.
You should consider making the Cancel button the default button for forms that support operations that can't be undone (such as delete). To do this, set both  **Cancel** and the **Default** property to **True**.

