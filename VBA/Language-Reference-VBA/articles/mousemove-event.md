---
title: MouseMove Event
keywords: fm20.chm2000180
f1_keywords:
- fm20.chm2000180
ms.prod: office
api_name:
- Office.MouseMove
ms.assetid: 0bbb767d-c113-2a65-7ca1-a3f98f4a3cce
ms.date: 06/08/2017
---


# MouseMove Event



Occurs when the user moves the mouse.
 **Syntax**
For MultiPage, TabStrip **Private Sub**_object_ _**MouseMove(**_index_**As Long**, **ByVal**_Button_**As fmButton**, **ByVal**_Shift_**As fmShiftState**, **ByVal**_X_**As Single**, **ByVal**_Y_**As Single)**
For other controls **Private Sub**_object_ _**MouseMove( ByVal**_Button_**As fmButton**, **ByVal**_Shift_**As fmShiftState**, **ByVal**_X_**As Single**, **ByVal**_Y_**As Single)**
The  **MouseMove** event syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object name.|
| _index_|Required. The index of the page or tab in a  **MultiPage** or **TabStrip** associated with this event.|
| _Button_|Required. An integer value that identifies the state of the mouse buttons.|
| _Shift_|Required. Specifies the state of SHIFT, CTRL, and ALT.|
| _X, Y_|Required. The horizontal or vertical position, measured in points, from the left or top edge of the control.|
 **Settings**
The  _index_ argument specifies which page or tab was clicked over. A _-_ 1 designates that the user did not click on any of the pages or tabs.
The settings for  _Button_ are:


|**Value**|**Description**|
|:-----|:-----|
|0|No button is pressed.|
|1|The left button is pressed.|
|2|The right button is pressed.|
|3|The right and left buttons are pressed.|
|4|The middle button is pressed.|
|5|The middle and left buttons are pressed.|
|6|The middle and right buttons are pressed.|
|7|All three buttons are pressed.|
The settings for  _Shift_ are:


|**Value**|**Description**|
|:-----|:-----|
|1|SHIFT was pressed.|
|2|CTRL was pressed.|
|3|SHIFT and CTRL were pressed.|
|4|ALT was pressed.|
|5|ALT and SHIFT were pressed.|
|6|ALT and CTRL were pressed.|
|7|ALT, SHIFT, and CTRL were pressed.|
You can identify individual keyboard modifiers by using the following constants:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| _fmShiftMask_|1|Mask to detect SHIFT.|
| _fmCtrlMask_|2|Mask to detect CTRL.|
| _fmAltMask_|4|Mask to detect ALT.|
 **Remarks**
The MouseMove event applies to forms, controls on a form, and labels.
MouseMove events are generated continually as the mouse pointer moves across objects. Unless another object has captured the mouse, an object recognizes a MouseMove event whenever the mouse position is within its borders.
Moving a form can also generate a MouseMove event even if the mouse is stationary. MouseMove events are generated when the form moves underneath the pointer. If a macro or event procedure moves a form in response to a MouseMove event, the event can continually generate (cascade) MouseMove events.
If two controls are very close together, and you move the mouse pointer quickly over the space between them, the MouseMove event might not occur for that space. In such cases, you might need to respond to the MouseMove event in both controls.
You can use the value returned in the  _Button_ argument to identify the state of the mouse buttons.
Use the  _Shift_ argument to identify the state of SHIFT, CTRL, and ALT when the MouseMove event occurred. For example, if both CTRL and ALT are pressed, the value of _Shift_ is 6.

 **Note**  You can use MouseDown and MouseUp event procedures to respond to events caused by pressing and releasing mouse buttons.


