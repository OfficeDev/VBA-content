---
title: Form.Unload Event (Access)
keywords: vbaac10.chm13644
f1_keywords:
- vbaac10.chm13644
ms.prod: access
api_name:
- Access.Form.Unload
ms.assetid: 13f1f7f4-9d69-128f-7e02-f3d3b99ec0f4
ms.date: 06/08/2017
---


# Form.Unload Event (Access)

The  **Unload** event occurs after a form is closed but before it's removed from the screen. When the form is reloaded, Microsoft Access redisplays the form and reinitializes the contents of all its controls.


## Syntax

 _expression_. **Unload**( ** _Cancel_**, )

 _expression_ A variable that represents a **Form** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required|**Integer**|Set to  **True** to cancel the **Unload** event.|

## Remarks

To run a macro or event procedure when these events occur, set the  **OnUnload** property to the name of the macro or to [Event Procedure].

The  **Unload** event is caused by user actions such as:


- Closing the form.
    
- Running the Close action in a macro.
    
- Quitting an application by right-clicking the application's taskbar button and then clicking  **Close**.
    
- Quitting Windows while an application is running.
    
By running a macro or an event procedure when a form's  **Unload** event occurs, you can verify that the form should be unloaded or specify actions that should take place when the form is unloaded. You can also open another form or display a dialog box requesting the user's name to make a log entry indicating who used the form.

When you close a form, the following events occur in this order:

**Unload** → **Deactivate** → **Close**

The  **Unload** event occurs before the **Close** event. The **Unload** event can be canceled, but the **Close** event can't.


 **Note**  When you create macros or event procedures for events related to the  **Unload** event, such as **Deactivate** and **LostFocus**, be sure that they don't conflict (for example, make sure you don't cause something to happen in one macro or procedure that is canceled in another) and that they don't cause cascading events.


## Example

This example prompts the user to verify that the form should close.

To try the example, add the following event procedure to a form. In Form view, close the form to display the dialog box, and then click Yes or No.




```vb
Private Sub Form_Unload(Cancel As Integer) 
 If MsgBox("Close form?", vbYesNo) = vbYes Then 
 Exit Sub 
 Else 
 Cancel = True 
 End If 
End Sub
```


## See also


#### Concepts


[Form Object](form-object-access.md)

