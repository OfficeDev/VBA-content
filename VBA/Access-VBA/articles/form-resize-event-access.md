---
title: Form.Resize Event (Access)
keywords: vbaac10.chm13643
f1_keywords:
- vbaac10.chm13643
ms.prod: access
api_name:
- Access.Form.Resize
ms.assetid: de57e9bf-e4fd-174e-4d56-9ea813ab92ce
ms.date: 06/08/2017
---


# Form.Resize Event (Access)

The  **Resize** event occurs when a form is opened and whenever the size of a form changes.


## Syntax

 _expression_. **Resize**

 _expression_ A variable that represents a **Form** object.


## Remarks

To run a macro or event procedure when this event occurs, set the  **[OnResize](form-onresize-property-access.md)** property to the name of the macro or to [Event Procedure].

This event occurs if you change the size of the form in a macro or event procedure— for example, if you use the MoveSize action in a macro to resize the form.

By running a macro or an event procedure when a  **Resize** event occurs, you can move or resize a control when the form it's on is resized. You can also use a **Resize** event to recalculate variables or reset properties that may depend on the size of the form.

When you first open a form, the following events occur in this order:

Open → Load → Resize → Activate → Current


 **Note**  You need to be careful if you use a  **MoveSize**, **Maximize**, **Minimize**, or **Restore** action (or the corresponding methods of the **DoCmd** object) in a **Resize** macro or event procedure. These actions can trigger a **Resize** event for the form, and thus cause a cascading event.


## Example

The following example shows how a  **Resize** event procedure can be used to repaint a form when it is maximized. When the user clicks a command button labeled "Maximize," the form is maximized and the **Resize** event is triggered.

To try the example, add the following event procedures to a form named Contacts that contains a command button named Maximize: 




```vb
Private Sub Maximize_Click() 
 DoCmd.Maximize 
End Sub 
 
Private Sub Form_Resize() 
 Forms!Contacts.Repaint 
End Sub
```


## See also


#### Concepts


[Form Object](form-object-access.md)

