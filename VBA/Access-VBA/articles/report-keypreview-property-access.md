---
title: Report.KeyPreview Property (Access)
keywords: vbaac10.chm13824
f1_keywords:
- vbaac10.chm13824
ms.prod: access
api_name:
- Access.Report.KeyPreview
ms.assetid: 49ca195d-bd9e-7a69-1891-455581bcf09a
ms.date: 06/08/2017
---


# Report.KeyPreview Property (Access)

You can use the  **KeyPreview** property to specify whether the report-level keyboard event procedures are invoked before a control's keyboard event procedures. Read/write **Boolean**.


## Syntax

 _expression_. **KeyPreview**

 _expression_ A variable that represents a **Report** object.


## Remarks

The  **KeyPreview** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|Yes|True|The report receives keyboard events first, and then the active control receives keyboard events. |
|No|False| (Default) Only the active control receives keyboard events.|
You can set the  **KeyPreview** property in any view.

You can use the  **KeyPreview** property to create a keyboard-handling procedure for a report. For example, when an application uses function keys, setting the **KeyPreview** property to **True** allows you to process keystrokes at the report level rather than writing code for each control that might receive keystroke events.

To handle keyboard events only at the report level and prevent controls from receiving keyboard events, set the  _KeyAscii_ argument to 0 in the report's **KeyPress** event procedure, and set the _KeyCode_ argument to 0 in the report's **KeyDown** and **KeyUp** event procedures.

If a report has no visible or enabled controls, it automatically receives all keyboard events.


## Example

In the following example, the  **KeyPreview** property is set to **True** in the report's **Load** event procedure. This causes the report to receive keyboard events before they are received by any control. The report's **KeyDown** event then checks the _KeyCode_ argument value to determine if the F2, F3, or F4 keys were pressed.


```vb
Private Sub Report_Load() 
 Me.KeyPreview = True 
End Sub 
 
Private Sub Report_KeyDown(KeyCode As Integer, Shift As Integer) 
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


[Report Object](report-object-access.md)

