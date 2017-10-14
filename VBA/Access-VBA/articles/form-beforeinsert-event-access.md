---
title: Form.BeforeInsert Event (Access)
keywords: vbaac10.chm13635
f1_keywords:
- vbaac10.chm13635
ms.prod: access
api_name:
- Access.Form.BeforeInsert
ms.assetid: de0f6b1a-fc11-4000-2c0c-b0ad9ccfccc2
ms.date: 06/08/2017
---


# Form.BeforeInsert Event (Access)

The BeforeInsert event occurs when the user types the first character in a new record, but before the record is actually created.


## Syntax

 _expression_. **BeforeInsert**( ** _Cancel_** )

 _expression_ A variable that represents a **Form** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required|**Integer**|The setting determines if the  **BeforeInsert** event occurs. Setting the _Cancel_ argument to **True** (?1) cancels the **BeforeInsert** event.|

## Remarks


 **Note**  Setting the value of a control by using a macro or Visual Basic doesn't trigger these events.

To run a macro or event procedure when these events occur, set the  **BeforeInsert** or **AfterInsert** property to the name of the macro or to [Event Procedure].

You can use an AfterInsert event procedure or macro to requery a recordset whenever a new record is added.

The BeforeInsert and AfterInsert events are similar to the  **BeforeUpdate** and **AfterUpdate** events. These events occur in the following order:

 **BeforeInsert** → **BeforeUpdate** → **AfterUpdate** → **AfterInsert**.

The following table summarizes the interaction between these events.



|**Event**|**Occurs when**|
|:-----|:-----|
|BeforeInsert|User types the first character in a new record.|
|BeforeUpdate|User updates the record.|
|AfterUpdate|Record is updated.|
|AfterInsert|Record updated is a new record.|
If the first character in a new record is typed into a text box or combo box, the  **BeforeInsert** event occurs before the **Change** event.


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

