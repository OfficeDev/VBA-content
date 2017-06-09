---
title: Attachment.Enter Event (Access)
keywords: vbaac10.chm14021
f1_keywords:
- vbaac10.chm14021
ms.prod: access
api_name:
- Access.Attachment.Enter
ms.assetid: 0ca691d8-aace-3240-c7c7-acfb69960f4a
ms.date: 06/08/2017
---


# Attachment.Enter Event (Access)

The  **Enter** event occurs before a control actually receives the focus from a control on the same form.


## Syntax

 _expression_. **Enter**

 _expression_ A variable that represents an **Attachment** object.


## Remarks

The  **Enter** event applies only to controls on a form, not controls on a report. This event does not apply to check boxes, option buttons, or toggle buttons in an option group. It applies only to the option group itself.

To run a macro or event procedure when these events occur, set the  **OnEnter** or **OnExit** property to the name of the macro or to [Event Procedure].

Because the  **Enter** event occurs before the focus moves to a particular control, you can use an **Enter** macro or event procedure to display instructions; for example, you could use a macro or event procedure to display a small form or message box identifying the type of data the control typically contains, or giving instructions on how to use the control.

The  **Enter** event occurs before the **GotFocus** event. The **Exit** event occurs before the **LostFocus** event.

Unlike the  **GotFocus** and **LostFocus** events, the **Enter** and **Exit** events don't occur when a form receives or loses the focus. For example, suppose you select a check box on a form, and then click a report. The **Enter** and **GotFocus** events occur when you select the check box. Only the **LostFocus** event occurs when you click the report. The **Exit** event doesn't occur (because the focus is moving to a different window). If you select the check box on the form again to bring it to the foreground, the **GotFocus** event occurs, but not the **Enter** event (because the control had the focus when the form was last active). The **Exit** event occurs only when you click another control on the form.

If you move the focus to a control on a form, and that control doesn't have the focus on that form, the  **Exit** and **LostFocus** events for the control that does have the focus on the form occur before the **Enter** and **GotFocus** events for the control to which you moved.

If you use the mouse to move the focus from a control on a main form to a control on a subform of that form (a control that doesn't already have the focus on the subform), the following events occur:

 **Exit** (for the control on the main form)

ß

 **LostFocus** (for the control on the main form)

ß

 **Enter** (for the subform control)

ß

 **Exit** (for the control on the subform that had the focus)

ß

 **LostFocus** (for the control on the subform that had the focus)

ß

 **Enter** (for the control on the subform to which the focus moved)

ß

 **GotFocus** (for the control on the subform to which the focus moved)

If the control you move to on the subform previously had the focus, neither its  **Enter** event nor its **GotFocus** event occurs, but the **Enter** event for the subform control does occur. If you move the focus from a control on a subform to a control on the main form, the **Exit** and **LostFocus** events for the control on the subform don't occur, just the **Exit** event for the subform control and the **Enter** and **GotFocus** events for the control on the main form.


 **Note**  You often use the mouse or a key such as TAB to move the focus to another control. This causes mouse or keyboard events to occur in addition to the events discussed in this topic.


## See also


#### Concepts


[Attachment Object](attachment-object-access.md)

