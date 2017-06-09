---
title: Report.KeyPress Event (Access)
keywords: vbaac10.chm13895
f1_keywords:
- vbaac10.chm13895
ms.prod: access
api_name:
- Access.Report.KeyPress
ms.assetid: 0c846367-a4b0-d716-dcc3-32c916e09dfb
ms.date: 06/08/2017
---


# Report.KeyPress Event (Access)

The  **KeyPress** event occurs when the user presses and releases a key or key combination that corresponds to an ANSI code while a report has the focus. This event also occurs if you send an ANSI keystroke to a report by using the SendKeys action in a macro or the **SendKeys** statement in Visual Basic.


## Syntax

 _expression_. **KeyPress**( ** _KeyAscii_**, )

 _expression_ A variable that represents a **Report** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _KeyAscii_|Required|**Integer**| Returns a numeric ANSI key code. The _KeyAscii_ argument is passed by reference; changing it sends a different character to the object. Setting the _KeyAscii_ argument to 0 cancels the keystroke so that the object doesn't recognize that a key was pressed.|

## Remarks

To run a macro or event procedure when this event occurs, set the  **OnKeyPress** property to the name of the macro or to [Event Procedure].

A report will also receive all keyboard events, even those that occur for controls, if you set the  **KeyPreview** property of the report to Yes. With this property setting, all keyboard events occur first for the report, and then for the control that has the focus. You can respond to specific keys pressed in the report, regardless of which control has the focus. For example, you may want the key combination CTRL+X to always perform the same action on a report.

If you press and hold down an ANSI key, the  **KeyDown** and **KeyPress** events alternate repeatedly ( **KeyDown**, **KeyPress**, **KeyDown**, **KeyPress**, and so on) until you release the key, and then the **KeyUp** event occurs.

A  **KeyPress** event can involve any printable keyboard character, the CTRL key combined with a character from the standard alphabet or a special character, and the ENTER or BACKSPACE key. You can use the **KeyDown** and **KeyUp** event procedures to handle any keystroke not recognized by the **KeyPress** event, such as function keys, navigation keys, and any combinations of these with keyboard modifiers (ALT, SHIFT, or CTRL keys). Unlike the **KeyDown** and **KeyUp** events, the **KeyPress** event doesn't indicate the physical state of the keyboard; instead, it indicates the ANSI character that corresponds to the pressed key or key combinations.

 **KeyPress** interprets the uppercase and lowercase version of each character as a separate key code and, therefore, as two separate characters.


 **Note**  The BACKSPACE key is part of the ANSI character set, but the DEL key is not. If you delete a character in a control by using the BACKSPACE key, you cause a  **KeyPress** event; if you use the DEL key, you don't.

The  **KeyDown** and **KeyPress** events occur when you press or send an ANSI key. The **KeyUp** event occurs after any event for a control caused by pressing or sending the key. If a keystroke causes the focus to move from one control to another control, the **KeyDown** event occurs for the first control, while the **KeyPress** and **KeyUp** events occur for the second control.


## See also


#### Concepts


[Report Object](report-object-access.md)

