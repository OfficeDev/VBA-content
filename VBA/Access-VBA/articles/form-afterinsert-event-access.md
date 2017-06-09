---
title: Form.AfterInsert Event (Access)
keywords: vbaac10.chm13636
f1_keywords:
- vbaac10.chm13636
ms.prod: access
api_name:
- Access.Form.AfterInsert
ms.assetid: 07140c13-ce7c-91f2-7451-d7f834653ef2
ms.date: 06/08/2017
---


# Form.AfterInsert Event (Access)

The  **AfterInsert** event occurs after a new record is added.


## Syntax

 _expression_. **AfterInsert**

 _expression_ A variable that represents a **Form** object.


### Return Value

nothing


## Remarks


 **Note**  Setting the value of a control by using a macro or Visual Basic doesn't trigger these events.

You can use an  **AfterInsert** event procedure or macro to requery a recordset whenever a new record is added.

To run a macro or event procedure when the  **AfterInsert** event occurs, set the **OnAfterInsert** property to the name of the macro or to [Event Procedure].


## Example

This example shows how you can use a  **BeforeInsert** event procedure to verify that the user wants to create a new record, and an **AfterInsert** event procedure to requery the record source for the Employees form after a record has been added.

To try the example, add the following event procedure to a form named Employees that is based on a table or query. Switch to form Datasheet view and try to insert a record.




```vb
Private Sub Form_BeforeInsert(Cancel As Integer) 
 If MsgBox("Insert new record here?", _ 
 vbOKCancel) = vbCancel Then 
 Cancel = True 
 End If 
End Sub 
 
Private Sub Form_AfterInsert() 
 Forms!Employees.Requery 
End Sub
```


## See also


#### Concepts


[Form Object](form-object-access.md)

