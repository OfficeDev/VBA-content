---
title: KeyDown, KeyUp Events
keywords: fm20.chm2000120
f1_keywords:
- fm20.chm2000120
ms.prod: office
ms.assetid: dde8140e-ebd7-2ad1-6585-65ffe65b3c22
ms.date: 06/08/2017
---


# KeyDown, KeyUp Events



Occur in sequence when a user presses and releases a key. KeyDown occurs when the user presses a key. KeyUp occurs when the user releases a key.
 **Syntax**
 **Private Sub**_object_ _**KeyDown( ByVal**_KeyCode_**As MSForms.ReturnInteger**, **ByVal**_Shift_**As fmShiftState)**
 **Private Sub**_object_ _**KeyUp( ByVal**_KeyCode_**As MSForms.ReturnInteger**, **ByVal**_Shift_**As fmShiftState)**
The  **KeyDown** and **KeyUp** event syntaxes have these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object name.|
| _KeyCode_|Required. An integer that represents the key code of the key that was pressed or released.|
| _Shift_|Required. The state of SHIFT, CTRL, and ALT.|
 **Settings**
The settings for  _Shift_ are:


|**Constant**|**Value**|**Description**|
|:-----|:-----|:-----|
| _fmShiftMask_|1|SHIFT was pressed.|
| _fmCtrlMask_|2|CTRL was pressed.|
| _fmAltMask_|4|ALT was pressed.|
 **Remarks**
The KeyDown event occurs when the user presses a key on a running form while that form or a control on it has the [focus](vbe-glossary.md). The KeyDown and KeyPress events alternate repeatedly until the user releases the key, at which time the KeyUp event occurs. The form or control with the focus receives all keystrokes. A form can have the focus only if it has no controls or all its visible controls are disabled.
These events also occur if you send a keystroke to a form or control using either the SendKeys action in a macro or the [SendKeys Statement](vbe-glossary.md) in Visual Basic.
The KeyDown and KeyUp events are typically used to recognize or distinguish between:


- Extended character keys, such as function keys.
    
- Navigation keys, such as HOME, END, PAGEUP, PAGEDOWN, UP ARROW, DOWN ARROW, RIGHT ARROW, LEFT ARROW, and TAB.
    
- Combinations of keys and standard keyboard modifiers (SHIFT, CTRL, or ALT).
    
- The numeric keypad and keyboard number keys.
    

The KeyDown and KeyUp events do not occur under the following circumstances:


- The user presses enter on a form with a command button whose  **Default** property is set to **True**.
    
- The user presses esc on a form with a command button whose  **Cancel** property is set to **True**.
    

The KeyDown and KeyPress events occur when you press or send an ANSI key. The KeyUp event occurs after any event for a control caused by pressing or sending the key. If a keystroke causes the focus to move from one control to another control, the KeyDown event occurs for the first control, while the KeyPress and KeyUp events occur for the second control.
The sequence of keyboard-related events is:


1. KeyDown
    
2. KeyPress
    
3. KeyUp
    


 **Note**  The KeyDown and KeyUp events apply only to forms and controls on a form. To interpret ANSI characters or to find out the ANSI character corresponding to the key pressed, use the KeyPress event.


