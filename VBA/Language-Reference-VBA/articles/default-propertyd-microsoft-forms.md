---
title: Default Property (Microsoft Forms)
keywords: fm20.chm2001070
f1_keywords:
- fm20.chm2001070
ms.prod: office
ms.assetid: d3c3a54c-5147-3ef5-545f-a1aece88d366
ms.date: 06/08/2017
---


# Default Property (Microsoft Forms)



Designates the default command button on a form.
 **Syntax**
 _object_. **Default** [= _Boolean_ ]
The  **Default** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _Boolean_|Optional. Whether the command button is the default.|
 **Settings**
The settings for  _Boolean_ are:


|**Value**|**Description**|
|:-----|:-----|
|**True**|The  **CommandButton** is the default button.|
|**False**|The  **CommandButton** is not the default button (default).|
 **Remarks**
A  **CommandButton** or an object that acts like a command button can be designated as the default command button. Only one object on a form can be the default command button. Setting the **Default** property to **True** for one object automatically sets it to **False** for all other objects on the form.
To choose the default command button on an active form, the user can click the button, or press ENTER when no other  **CommandButton** has the[focus](vbe-glossary.md). Pressing ENTER when no other  **CommandButton** has the focus also initiates the KeyUp event for the default command button.
 **Default** is provided for[OLE container controls](glossary-vba.md) (Windows only) that specifically act like **CommandButton** controls.
 **Tip** You should consider making the Cancel button the default button for forms that support operations that can't be undone (such as delete). To do this, set both **Default** and the **Cancel** property to **True**.

