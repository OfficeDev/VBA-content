---
title: Form.Open Event (Access)
keywords: vbaac10.chm13642
f1_keywords:
- vbaac10.chm13642
ms.prod: access
api_name:
- Access.Form.Open
ms.assetid: 8638e6d9-29af-a007-44f5-9bada14adb29
ms.date: 06/08/2017
---


# Form.Open Event (Access)

The  **Open** event occurs when a form is opened, but before the first record is displayed.


## Syntax

 _expression_. **Open**( ** _Cancel_**, )

 _expression_ A variable that represents a **Form** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required|**Integer**|The setting determines if the opening of the form or report occurs. Setting the Cancel argument to True (?1) cancels the opening of the form or report.|

## Remarks

By running a macro or an event procedure when a form's  **Open** event occurs, you can close another window or move the focus to a particular control on a form. You can also run a macro or an event procedure that asks for information needed before the form or report is opened or printed.

For example, an  **Open** macro or event procedure can open a custom dialog box in which the user enters the criteria to filter the set of records to display on a form or the date range to include for a report.

The  **Open** event doesn't occur when you activate a form that's already open ? for example, when you switch to the form from another window in Microsoft Access or use the OpenForm action in a macro to bring the open form to the top. However, the **Activate** event does occur in these situations.

When you open a form based on an underlying query, Microsoft Access runs the underlying query for the form before it runs the Open macro or event procedure.

If your application can have more than one form loaded at a time, use the  **Activate** and **Deactivate** events instead of the Open event to display and hide custom toolbars when the focus moves to a different form.

The Open event occurs before the  **Load** event, which is triggered when a form is opened and its records are displayed.

When you first open a form, the following events occur in this order:

 **Open** → **Load** → **Resize** → **Activate** → **Current**

The Close event occurs after the  **Unload** event, which is triggered after the form is closed but before it is removed from the screen.

When you close a form, the following events occur in this order:

 **Unload** → **Deactivate** → **Close**

When the Close event occurs, you can open another window or request the user's name to make a log entry indicating who used the form or report.

If you're trying to decide whether to use the  **Open** or Load event for your macro or event procedure, one significant difference is that the **Open** event can be canceled, but the **Load** event can't. For example, if you're dynamically building a record source for a form in an event procedure for the form's **Open** event, you can cancel opening the form if there are no records to display. Similarly, the **Unload** event can be canceled, but the **Close** event can't.


## Example

The following example shows how you can cancel the opening of a form when the user clicks a No button. A message box prompts the user to enter order details. If the user clicks No, the Order Details form isn't opened.

To try the example, add the following event procedure to a form.




```vb
Private Sub Form_Open(Cancel As Integer) 
 Dim intReturn As Integer 
 intReturn = MsgBox("Enter order details now?", vbYesNo) 
 Select Case intReturn 
 Case vbYes 
 ' Open Order Details form. 
 DoCmd.OpenForm "Order Details" 
 Case vbNo 
 MsgBox "Remember to enter order details by 5 P.M." 
 Cancel = True ' Cancel Open event. 
 End Select 
End Sub
```


## See also


#### Concepts


[Form Object](form-object-access.md)

