---
title: ControlTipText Property
keywords: fm20.chm2000990
f1_keywords:
- fm20.chm2000990
ms.prod: office
api_name:
- Office.ControlTipText
ms.assetid: 879e93e6-7646-1707-ff43-1b66882da4cd
ms.date: 06/08/2017
---


# ControlTipText Property



Specifies text that appears when the user briefly holds the mouse pointer over a control without clicking.
 **Syntax**
 _object_. **ControlTipText** [= _String_ ]
The  **ControlTipText** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _String_|Optional. The text that appears when the user holds the mouse pointer over a control.|
 **Remarks**
The  **ControlTipText** property lets you give users tips about a control in a running form. The property can be set during[design time](vbe-glossary.md) but only appears by the control during[run time](vbe-glossary.md).
The default value of  **ControlTipText** is an empty string. When the value of **ControlTipText** is set to an empty string, no tip is available for that control.

