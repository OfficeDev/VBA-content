---
title: KeyPress Event
keywords: fm20.chm2000130
f1_keywords:
- fm20.chm2000130
ms.prod: office
api_name:
- Office.KeyPress
ms.assetid: b77da9a6-a87c-a44c-ab44-02495af3fa5e
ms.date: 06/08/2017
---


# KeyPress Event



Occurs when the user presses an ANSI key.
 **Syntax**
 **Private Sub**_object_ _**KeyPress( ByVal**_KeyANSI_**As MSForms.ReturnInteger)**
The  **KeyPress** event syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _object_|Required. A valid object.|
| _KeyANSI_|Required. An integer value that represents a standard numeric ANSI key code.|
 **Remarks**
The KeyPress event occurs when the user presses a key that produces a typeable character (an ANSI key) on a running form while the form or a control on it has the [focus](vbe-glossary.md). The event can occur either before or after the key is released. This event also occurs if you send an ANSI keystroke to a form or control using either the SendKeys action in a macro or the [SendKeys Statement](vbe-glossary.md) in Visual Basic.
A KeyPress event can occur when any of the following keys are pressed:


- Any printable keyboard character.
    
- CTRL combined with a character from the standard alphabet.
    
- CTRL combined with any special character.
    
- BACKSPACE.
    
- ESC.
    

A KeyPress event does not occur under the following conditions:


- Pressing TAB.
    
- Pressing ENTER.
    
- Pressing an arrow key.
    
- When a keystroke causes the focus to move from one control to another.
    


 **Note**  BACKSPACE is part of the [ANSI character set](vbe-glossary.md), but DELETE is not. Deleting a character in a control using BACKSPACE causes a KeyPress event; deleting a character using DELETE doesn't.

When a user holds down a key that produces an ANSI keycode, the KeyDown and KeyPress events alternate repeatedly. When the user releases the key, the KeyUp event occurs. The form or control with the focus receives all keystrokes. A form can have the focus only if it has no controls, or if all its visible controls are disabled.
The default action for the KeyPress event is to process the event code that corresponds to the key that was pressed.  _KeyANSI_ indicates the ANSI character that corresponds to the pressed key or key combination. The KeyPress event interprets the uppercase and lowercase of each character as separate key codes and, therefore, as two separate characters.
To respond to the physical state of the keyboard, or to handle keystrokes not recognized by the KeyPress event, such as function keys, navigation keys, and any combinations of these with keyboard modifiers (ALT, SHIFT, or CTRL), use the KeyDown and KeyUp event procedures.
The sequence of keyboard-related events is:


1. KeyDown
    
2. KeyPress
    
3. KeyUp
    


