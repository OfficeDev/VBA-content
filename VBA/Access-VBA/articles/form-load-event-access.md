---
title: Form.Load Event (Access)
keywords: vbaac10.chm13633
f1_keywords:
- vbaac10.chm13633
ms.prod: access
api_name:
- Access.Form.Load
ms.assetid: a7547066-e1eb-6cdc-a170-2ee222081720
ms.date: 06/08/2017
---


# Form.Load Event (Access)

Occurs when a form is opened and its records are displayed.


## Syntax

 _expression_. **Load**

 _expression_ A variable that represents a **Form** object.


## Remarks

To run a macro or event procedure when these events occur, set the  **OnLoad** property to the name of the macro or to [Event Procedure].

The  **Load** event is caused by user actions such as:


- Starting an application. 
    
- Opening a form by clicking Open in the Database window. 
    
- Running the OpenForm action in a macro.
    
By running a macro or an event procedure when a form's  **Load** event occurs, you can specify default settings for controls, or display calculated data that depends on the data in the form's records.

By running a macro or an event procedure when a form's  **Unload** event occurs, you can verify that the form should be unloaded or specify actions that should take place when the form is unloaded. You can also open another form or display a dialog box requesting the user's name to make a log entry indicating who used the form.

When you first open a form, the following events occur in this order:

Open → Load → Resize → Activate → Current

If you're trying to decide whether to use the  **Open** or **Load** event for your macro or event procedure, one significant difference is that the **Open** event can be canceled, but the **Load** event can't. For example, if you're dynamically building a record source for a form in an event procedure for the form's **Open** event, you can cancel opening the form if there are no records to display.

When you close a form, the following events occur in this order:

Unload → Deactivate → Close

The  **Unload** event occurs before the **Close** event. The **Unload** event can be canceled, but the **Close** event can't.


 **Note**  When you create macros or event procedures for events related to the  **Load** event, such as **Activate** and **GotFocus**, be sure that they don't conflict (for example, make sure you don't cause something to happen in one macro or procedure that is canceled in another) and that they don't cause cascading events.


## Example

The following example displays the current date in the form's caption when the form is loaded.

To try the example, add the following event procedure to a form:




```vb
Private Sub Form_Load() 
 Me.Caption = Date 
End Sub
```


## See also


#### Concepts


[Form Object](form-object-access.md)

