---
title: ComboBox.Exit Event (Access)
keywords: vbaac10.chm14216
f1_keywords:
- vbaac10.chm14216
ms.prod: access
api_name:
- Access.ComboBox.Exit
ms.assetid: 47f37eb3-c0c1-457f-31ec-3b33b02ba986
ms.date: 06/08/2017
---


# ComboBox.Exit Event (Access)

The  **Exit** event occurs just before a control loses the focus to another control on the same form or report.


## Syntax

 _expression_. **Exit**( ** _Cancel_** )

 _expression_ A variable that represents a **ComboBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required|**Integer**|The setting determines if the  **Exit** event occurs. Setting the _Cancel_ argument to **True** (?1) cancels the **Exit** event.|

## Remarks

This event does not apply to check boxes, option buttons, or toggle buttons in an option group. It applies only to the option group itself.

To run a macro or event procedure when this event occurs, set the  **OnExit** property to the name of the macro or to [Event Procedure].

Because the  **Enter** event occurs before the focus moves to a particular control, you can use an **Enter** macro or event procedure to display instructions; for example, you could use a macro or event procedure to display a small form or message box identifying the type of data the control typically contains, or giving instructions on how to use the control.

The  **Exit** event occurs before the **LostFocus** event.

Unlike the  **LostFocus** event, the **Exit** event does not occur when a form loses the focus. For example, suppose you select a check box on a form, and then click a report. The **Enter** and **GotFocus** events occur when you select the check box. Only the **LostFocus** event occurs when you click the report. The **Exit** event doesn't occur (because the focus is moving to a different window). If you select the check box on the form again to bring it to the foreground, the **GotFocus** event occurs, but not the **Enter** event (because the control had the focus when the form was last active). The **Exit** event occurs only when you click another control on the form.

If you move the focus to a control on a form, and that control doesn't have the focus on that form, the  **Exit** and **LostFocus** events for the control that does have the focus on the form occur before the **Enter** and **GotFocus** events for the control you moved to.

If you use the mouse to move the focus from a control on a main form to a control on a subform of that form (a control that doesn't already have the focus on the subform), the following events occur:

 **Exit** (for the control on the main form)

?

 **LostFocus** (for the control on the main form)

?

 **Enter** (for the subform control)

?

 **Exit** (for the control on the subform that had the focus)

?

 **LostFocus** (for the control on the subform that had the focus)

?

 **Enter** (for the control on the subform that the focus moved to)

?

 **GotFocus** (for the control on the subform that the focus moved to)

If the control you move to on the subform previously had the focus, neither its  **Enter** event nor its **GotFocus** event occurs, but the **Enter** event for the subform control does occur. If you move the focus from a control on a subform to a control on the main form, the **Exit** and **LostFocus** events for the control on the subform don't occur, just the **Exit** event for the subform control and the **Enter** and **GotFocus** events for the control on the main form.


 **Note**  You often use the mouse or a key such as TAB to move the focus to another control. This causes mouse or keyboard events to occur in addition to the events discussed in this topic.


## Example

In the following example, two event procedures are attached to the LastName text box. The  **Enter** event procedure displays a message specifying what type of data the user can enter in the text box. The **Exit** event procedure displays a dialog box asking the user if changes should be saved before the focus moves to another control. If the user clicks the Cancel button, the _Cancel_ argument is set to **True** (?1), which moves the focus to the text box without saving changes. If the user chooses the OK button, the changes are saved, and the focus moves to another control.

To try the example, add the following event procedure to a form that contains a text box named LastName.




```vb
Private Sub LastName_Enter() 
 MsgBox "Enter your last name." 
End Sub 
 
Private Sub LastName_Exit(Cancel As Integer) 
 Dim strMsg As String 
 
 strMsg = "You entered '" &; Me!LastName _ 
 &; "' as your last name." &; _ 
 vbCrLf &; "Is this correct?" 
 If MsgBox(strMsg, vbYesNo) = vbNo Then 
 Cancel = True ' Cancel exit. 
 Else 
 Exit Sub ' Save changes and exit. 
 End If 
End Sub
```


## See also


#### Concepts


[ComboBox Object](combobox-object-access.md)

