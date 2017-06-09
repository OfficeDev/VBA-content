---
title: MousePointer Property
keywords: fm20.chm2001550
f1_keywords:
- fm20.chm2001550
ms.prod: office
api_name:
- Office.MousePointer
ms.assetid: ae574d87-e218-4d03-d423-0192768e82dc
ms.date: 06/08/2017
---


# MousePointer Property



Specifies the type of pointer displayed when the user positions the mouse over a particular object.
 **Syntax**
 _object_. **MousePointer** [= _fmMousePointer_ ]
The  **MousePointer** property syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _fmMousePointer_|Optional. The shape you want for the mouse pointer.|
 **Settings**
The settings for  _fmMousePointer_ are:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| _fmMousePointerDefault_|0|Standard pointer. The image is determined by the object (default).|
| _fmMousePointerArrow_|1|Arrow.|
| _fmMousePointerCross_|2|Cross-hair pointer.|
| _fmMousePointerIBeam_|3|I-beam.|
| _fmMousePointerSizeNESW_|6|Double arrow pointing northeast and southwest.|
| _fmMousePointerSizeNS_|7|Double arrow pointing north and south.|
| _fmMousePointerSizeNWSE_|8|Double arrow pointing northwest and southeast.|
| _fmMousePointerSizeWE_|9|Double arrow pointing west and east.|
| _fmMousePointerUpArrow_|10|Up arrow.|
| _fmMousePointerHourglass_|11|Hourglass.|
| _fmMousePointerNoDrop_|12|"Not" symbol (circle with a diagonal line) on top of the object being dragged. Indicates an invalid drop target.|
| _fmMousePointerAppStarting_|13|Arrow with an hourglass.|
| _fmMousePointerHelp_|14|Arrow with a question mark.|
| _fmMousePointerSizeAll_|15|Size all cursor (arrows pointing north, south, east, and west).|
| _fmMousePointerCustom_|99|Uses the icon specified by the  **MouseIcon** property.|
 **Remarks**
Use the  **MousePointer** property when you want to indicate changes in functionality as the mouse pointer passes over controls on a form. For example, the hourglass setting (11) is useful to indicate that the user must wait for a process or operation to finish.
Some icons vary depending on system settings, such as the icons associated with desktop themes.

