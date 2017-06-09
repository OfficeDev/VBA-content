---
title: Caption Property (Microsoft Forms)
keywords: fm20.chm916518
f1_keywords:
- fm20.chm916518
ms.prod: office
ms.assetid: d2303a41-d557-032c-c195-febde9029f8a
ms.date: 06/08/2017
---


# Caption Property (Microsoft Forms)



Descriptive text that appears on an object to identify or describe it.
 **Syntax**
 _object_. **Caption** [= _String_ ]
The  **Caption** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _String_|Optional. A string expression that evaluates to the text displayed as the caption.|
 **Settings**
The default setting for a control is a unique name based on the type of control. For example, CommandButton1 is the default caption for the first command button in a form.
 **Remarks**
The text identifies or describes the object with which it is associated. For buttons and labels, the  **Caption** property specifies the text that appears in the control. For **Page** and **Tab** objects, it specifies the text that appears on the tab.
If a control's caption is too long, the caption is truncated. If a form's caption is too long for the title bar, the title is displayed with an ellipsis.
The  **ForeColor** property of the control determines the color of the text in the caption.

 **Tip**  If a control has both the  **Caption** and **AutoSize** properties, setting **AutoSize** to **True** automatically adjusts the size of the control to frame the entire caption.


