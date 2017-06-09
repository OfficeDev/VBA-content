---
title: Attachment.KeyDown Event (Access)
keywords: vbaac10.chm14032
f1_keywords:
- vbaac10.chm14032
ms.prod: access
api_name:
- Access.Attachment.KeyDown
ms.assetid: 91a000e2-0a4e-4dd0-2715-b1987eb7212a
ms.date: 06/08/2017
---


# Attachment.KeyDown Event (Access)

The  **KeyDown** event occurs when the user presses a key while a form or control has the focus. This event also occurs if you send a keystroke to a form or control by using the SendKeys action in a macro or the **SendKeys** statement in Visual Basic.


## Syntax

 _expression_. **KeyDown**( ** _KeyCode_**, ** _Shift_** )

 _expression_ A variable that represents an **Attachment** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _KeyCode_|Required|**Integer**|A key code, such as  **vbKeyF1** (the F1 key) or **vbKeyHome** (the HOME key). To specify key codes, use the intrinsic constants shown in the Object Browser. You can prevent an object from receiving a keystroke by setting KeyCode to 0.|
| _Shift_|Required|**Integer**|The state of the SHIFT, CTRL, and ALT keys at the time of the event. If you need to test for the Shift argument, you can use one of the following intrinsic constants as bit masks:
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p><b>acShiftMask</b>  The bit mask for the SHIFT key.  
  </p></li><li><p><b>acCtrlMask</b>  The bit mask for the CTRL key.  
  </p></li><li><p><b>acAltMask</b>  The bit mask for the ALT key.  
</p></li></ul>|

## Remarks


 **Note**  The  **KeyDown** event applies only to forms and controls on a form, not controls on a report.

To run a macro or event procedure when these events occur, set the  **OnKeyDown** property to the name of the macro or to [Event Procedure].

For both events, the object with the focus receives all keystrokes. A form can have the focus only if it has no controls or all its visible controls are disabled.

A form will also receive all keyboard events, even those that occur for controls, if you set the  **KeyPreview** property of the form to Yes. With this property setting, all keyboard events occur first for the form, and then for the control that has the focus. You can respond to specific keys pressed in the form, regardless of which control has the focus. For example, you may want the key combination CTRL+X to always perform the same action on a form.

If you press and hold down a key, the  **KeyDown** and **KeyPress** events alternate repeatedly ( **KeyDown**, **KeyPress**, **KeyDown**, **KeyPress**, and so on) until you release the key, and then the **KeyUp** event occurs.

Although the  **KeyDown** event occurs when most keys are pressed, it is typically used to recognize or distinguish between:


- Extended character keys, such as function keys.
    
- Navigation keys, such as HOME, END, PAGE UP, PAGE DOWN, UP ARROW, DOWN ARROW, RIGHT ARROW, LEFT ARROW, and TAB.
    
- Combinations of keys and standard keyboard modifiers (SHIFT, CTRL, or ALT keys).
    
- The numeric keypad and keyboard number keys.
    
The  **KeyDown** event does not occur when you press:


- The ENTER key if the form has a command button for which the  **Default** property is set to Yes.
    
- The ESC key if the form has a command button for which the  **Cancel** property is set to Yes.
    
The  **KeyDown** event occurs when you press or send an ANSI key. The **KeyUp** event occurs after any event for a control caused by pressing or sending the key. If a keystroke causes the focus to move from one control to another control, the **KeyDown** event occurs for the first control, while the **KeyPress** and **KeyUp** events occur for the second control.

To find out the ANSI character corresponding to the key pressed, use the  **KeyPress** event.

If a modal dialog box is displayed as a result of pressing or sending a key, the  **KeyDown** and **KeyPress** events occur, but the **KeyUp** event doesn't occur.


## See also


#### Concepts


[Attachment Object](attachment-object-access.md)

