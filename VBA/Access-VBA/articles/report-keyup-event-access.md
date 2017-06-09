---
title: Report.KeyUp Event (Access)
keywords: vbaac10.chm13896
f1_keywords:
- vbaac10.chm13896
ms.prod: access
api_name:
- Access.Report.KeyUp
ms.assetid: 5561cbab-b6bd-ab4e-83a6-fbf7ec9272d1
ms.date: 06/08/2017
---


# Report.KeyUp Event (Access)

The  **KeyUp** event occurs when the user releases a key while a report has the focus. This event also occurs if you send a keystroke to a report by using the SendKeys action in a macro or the **SendKeys** statement in Visual Basic.


## Syntax

 _expression_. **KeyUp**( ** _KeyCode_**, ** _Shift_** )

 _expression_ A variable that represents a **Report** object.


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

To run a macro or event procedure when these events occur, set the  **OnKeyUp** property to the name of the macro or to [Event Procedure].

A report will also receive all keyboard events, even those that occur for controls, if you set the  **KeyPreview** property of the report to Yes. With this property setting, all keyboard events occur first for the report, and then for the control that has the focus. You can respond to specific keys pressed in the report, regardless of which control has the focus. For example, you may want the key combination CTRL+X to always perform the same action on a report.

If you press and hold down a key, the  **KeyDown** and **KeyPress** events alternate repeatedly ( **KeyDown**, **KeyPress**, **KeyDown**, **KeyPress**, and so on) until you release the key, and then the **KeyUp** event occurs.

Although the  **KeyUp** event occurs when most keys are pressed, it is typically used to recognize or distinguish between:


- Extended character keys, such as function keys.
    
- Navigation keys, such as HOME, END, PAGE UP, PAGE DOWN, UP ARROW, DOWN ARROW, RIGHT ARROW, LEFT ARROW, and TAB.
    
- Combinations of keys and standard keyboard modifiers (SHIFT, CTRL, or ALT keys).
    
- The numeric keypad and keyboard number keys.
    
To find out the ANSI character corresponding to the key pressed, use the  **KeyPress** event.

If a modal dialog box is displayed as a result of pressing or sending a key, the  **KeyDown** and **KeyPress** events occur, but the **KeyUp** event doesn't occur.


## See also


#### Concepts


[Report Object](report-object-access.md)

