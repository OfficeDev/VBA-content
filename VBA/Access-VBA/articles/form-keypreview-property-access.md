---
title: Form.KeyPreview Property (Access)
keywords: vbaac10.chm13457
f1_keywords:
- vbaac10.chm13457
ms.prod: access
api_name:
- Access.Form.KeyPreview
ms.assetid: f9153ec0-8b6e-60d5-8541-100e2ad1705e
ms.date: 06/08/2017
---


# Form.KeyPreview Property (Access)

You can use the  **KeyPreview** property to specify whether the form-level keyboard event procedures are invoked before a control's keyboard event procedures. Read/write **Boolean**.


## Syntax

 _expression_. **KeyPreview**

 _expression_ A variable that represents a **Form** object.


## Remarks

The KeyPreview property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Yes|True|The form receives keyboard events first, then the active control receives keyboard events. |
|No|False| (Default) Only the active control receives keyboard events.|
You can set the  **KeyPreview** property in any view

You can use the  **KeyPreview** property to create a keyboard-handling procedure for a form. For example, when an application uses function keys, setting the **KeyPreview** property to **True** allows you to process keystrokes at the form level rather than writing code for each control that might receive keystroke events.

To handle keyboard events only at the form level and prevent controls from receiving keyboard events, set the  _KeyAscii_ argument to 0 in the form's **KeyPress** event procedure, and set the _KeyCode_ argument to 0 in the form's **KeyDown** and **KeyUp** event procedures.

If a form has no visible or enabled controls, it automatically receives all keyboard events.


## Example

In the following example, the  **KeyPreview** property is set to **True** in the form's **Load** event procedure. This causes the form to receive keyboard events before they are received by any control. The form **KeyDown** event then checks the _KeyCode_ argument value to determine if the F2, F3, or F4 keys were pushed.


```vb
Private Sub Form_Load() 
 Me.KeyPreview = True 
End Sub 
 
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer) 
 Select Case KeyCode 
 Case vbKeyF2 
 ' Process F2 key events. 
 Case vbKeyF3 
 ' Process F3 key events. 
 Case vbKeyF4 
 ' Process F4 key events. 
 Case Else 
 End Select 
End Sub
```


## See also


#### Concepts


[Form Object](form-object-access.md)

