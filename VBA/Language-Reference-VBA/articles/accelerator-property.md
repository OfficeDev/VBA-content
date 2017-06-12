---
title: Accelerator Property
keywords: fm20.chm2000690
f1_keywords:
- fm20.chm2000690
ms.prod: office
api_name:
- Office.Accelerator
ms.assetid: d9183848-4638-745b-e3f4-b076493d3668
ms.date: 06/08/2017
---


# Accelerator Property



Sets or retrieves the [accelerator key](glossary-vba.md) for a control.
 **Syntax**
 _object_. **Accelerator** [= _String_ ]
The  **Accelerator** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _String_|Optional. The character to use as the accelerator key.|
 **Remarks**
To designate an accelerator key, enter a single character for the  **Accelerator** property. You can set **Accelerator** in the control's property sheet or in code. If the value of this property contains more than one character, the first character in the string becomes the value of **Accelerator**.
When an accelerator key is used, there is no visual feedback (other than [focus](vbe-glossary.md)) to indicate that the control initiated the Click event. For example, if the accelerator key applies to a  **CommandButton**, the user will not see the button pressed in the interface. The button receives the focus, however, when the user presses the accelerator key.
If the accelerator applies to a  **Label**, the control following the **Label** in the[tab order](vbe-glossary.md), rather than the  **Label** itself, receives the focus.

