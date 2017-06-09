---
title: NavigationButton.GotFocus Event (Access)
keywords: vbaac10.chm14080
f1_keywords:
- vbaac10.chm14080
ms.prod: access
api_name:
- Access.NavigationButton.GotFocus
ms.assetid: 3adf6a7e-34d5-e1ce-e621-8662153156e9
ms.date: 06/08/2017
---


# NavigationButton.GotFocus Event (Access)

The  **GotFocus** event occurs when the specified object receives the focus.


## Syntax

 _expression_. **GotFocus**

 _expression_ A variable that represents a **NavigationButton** object.


### Return Value

nothing


## Remarks

To run a macro or event procedure when these events occur, set the  **OnGotFocus** property to the name of the macro or to [Event Procedure].

These events occur when the focus moves in response to a user action, such as pressing the TAB key or clicking the object, or when you use the  **SetFocus** method in Visual Basic or the SelectObject, GoToRecord, GoToControl, or GoToPage action in a macro.

A control can receive the focus only if its  **Visible** and **Enabled** properties are set to Yes. A form can receive the focus only if it has no controls or if all visible controls are disabled. If a form contains any visible, enabled controls, the GotFocus event for the form doesn't occur.

You can specify what happens when a form or control receives the focus by running a macro or an event procedure when the  **GotFocus** event occurs. For example, by attaching a **GotFocus** event procedure to each control on a form, you can guide the user through your application by displaying brief instructions or messages in a text box. You can also provide visual cues by enabling, disabling, or displaying controls that depend on the control with the focus.


 **Note**  To customize the order in which the focus moves from control to control on a form when you press the TAB key, set the tab order or specify access keys for the controls.

The  **GotFocus** event differs from the **Enter** event in that the **GotFocus** event occurs every time a control receives the focus. For example, suppose the user clicks a check box on a form, then clicks a report, and finally clicks the check box on the form to bring it to the foreground. The **GotFocus** event occurs both times the check box receives the focus. In contrast, the **Enter** event occurs only the first time the user clicks the check box. The **GotFocus** event occurs after the **Enter** event.

If you move the focus to a control on a form, and that control doesn't have the focus on that form, the  **Exit** and **LostFocus** events for the control that does have the focus on the form occur before the **Enter** and **GotFocus** events for the control you moved to.

If you use the mouse to move the focus from a control on a main form to a control on a subform of that form, the following events occur:

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

 **Enter** (for the control on the subform that the focus moved to)

ß

 **GotFocus** (for the control on the subform that the focus moved to)

If the control you move to on the subform previously had the focus, neither its  **Enter** event nor its **GotFocus** event occurs, but the **Enter** event for the subform control does occur. If you move the focus from a control on a subform to a control on the main form, the **Exit** and **LostFocus** events for the control on the subform don't occur, just the **Exit** event for the subform control and the **Enter** and **GotFocus** events for the control on the main form.


 **Note**  You often use the mouse or a key such as TAB to move the focus to another control. This causes mouse or keyboard events to occur in addition to the events discussed in this topic.

When you switch between two open forms, the Deactivate event occurs for the first form, and the  **Activate** event occurs for the second form. If the forms contain no visible, enabled controls, the **LostFocus** event occurs for the first form before the **Deactivate** event, and the **GotFocus** event occurs for the second form after the **Activate** event.


## Example

The following example displays a message in a label when the focus moves to an option button.

To try the example, add the following event procedures to a form named Contacts that contains an option button named OptionYes and a label named LabelYes.




```vb
Private Sub OptionYes_GotFocus() 
 Me!LabelYes.Caption = "Option button 'Yes' has the focus." 
End Sub 
 
Private Sub OptionYes_LostFocus() 
 Me!LabelYes.Caption = "" ' Clear caption. 
End Sub
```


## See also


#### Concepts


[NavigationButton Object](navigationbutton-object-access.md)

