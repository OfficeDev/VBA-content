---
title: MouseDown, MouseUp Events
keywords: fm20.chm5224947
f1_keywords:
- fm20.chm5224947
ms.prod: office
ms.assetid: 760c2492-4a33-8d17-eeef-e52da662d4c4
ms.date: 06/08/2017
---


# MouseDown, MouseUp Events



Occur when the user clicks a mouse button. MouseDown occurs when the user presses the mouse button; MouseUp occurs when the user releases the mouse button.
 **Syntax**
For MultiPage, TabStrip **Private Sub**_object_ _**MouseDown(**_index_**As Long**, **ByVal**_Button_**As fmButton**, **ByVal**_Shift_**As fmShiftState**, **ByVal**_X_**As Single**, **ByVal**_Y_**As Single)** **Private Sub**_object_ _**MouseUp(**_index_**As Long**, **ByVal**_Button_**As fmButton**, **ByVal**_Shift_**As fmShiftState**, **ByVal**_X_**As Single**, **ByVal**_Y_**As Single)**
For other controls **Private Sub**_object_ _**MouseDown( ByVal**_Button_**As fmButton**, **ByVal**_Shift_**As fmShiftState**, **ByVal**_X_**As Single**, **ByVal**_Y_**As Single)** **Private Sub**_object_ _**MouseUp( ByVal**_Button_**As fmButton**, **ByVal**_Shift_**As fmShiftState**, **ByVal**_X_**As Single**, **ByVal**_Y_**As Single)**
The  **MouseDown** and **MouseUp** event syntaxes have these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _index_|Required. The index of the page or tab in a  **MultiPage** or **TabStrip** with the specified event.|
| _Button_|Required. An integer value that identifies which mouse button caused the event.|
| _Shift_|Required. The state of SHIFT, CTRL, and ALT.|
| _X, Y_|Required. The horizontal or vertical position, in points, from the left or top edge of the form,  **Frame**, or **Page**.|
 **Settings**
The settings for  _Button_ are:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| _fmButtonLeft_|1|The left button was pressed.|
| _fmButtonRight_|2|The right button was pressed.|
| _fmButtonMiddle_|4|The middle button was pressed.|
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
For a  **MultiPage**, the MouseDown event occurs when the user presses a mouse button over the control.
For a  **TabStrip**, the index argument identifies the tab where the user clicked. An index of -1 indicates the user did not click a tab. For example, if there are no tabs in the upper right corner of the control, clicking in the upper right corner sets the index to -1.
For a form, the user can generate MouseDown and MouseUp events by pressing and releasing a mouse button in a blank area, record selector, or scroll bar on the form.
The sequence of mouse-related events is:


1. MouseDown
    
2. MouseUp
    
3. Click
    
4. DblClick
    
5. MouseUp
    

MouseDown or MouseUp event procedures specify actions that occur when a mouse button is pressed or released. MouseDown and MouseUp events enable you to distinguish between the left, right, and middle mouse buttons. You can also write code for mouse-keyboard combinations that use the SHIFT, CTRL, and ALT keyboard modifiers.
If a mouse button is pressed while the pointer is over a form or control, that object "captures" the mouse and receives all mouse events up to and including the last MouseUp event. This implies that the  _X_, _Y_ mouse-pointer coordinates returned by a mouse event may not always be within the boundaries of the object that receives them.
If mouse buttons are pressed in succession, the object that captures the mouse receives all successive mouse events until all buttons are released.
Use the  _Shift_ argument to identify the state of SHIFT, CTRL, and ALT when the MouseDown or MouseUp event occurred. For example, if both CTRL and ALT are pressed, the value of _Shift_ is 6.

