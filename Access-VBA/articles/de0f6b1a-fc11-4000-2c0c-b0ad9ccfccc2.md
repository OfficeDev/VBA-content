
# Form.BeforeInsert Event (Access)

 **Last modified:** July 28, 2015

 **In this article**
 [Syntax](#sectionSection0)
 [Remarks](#sectionSection1)
 [Example](#sectionSection2)


The BeforeInsert event occurs when the user types the first character in a new record, but before the record is actually created.


## Syntax
<a name="sectionSection0"> </a>

 _expression_. **BeforeInsert**( **_Cancel_** )

 _expression_A variable that represents a  **Form** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Cancel|Required| **Integer**|The setting determines if the  **BeforeInsert** event occurs. Setting theCancel argument to **True** (-1) cancels the **BeforeInsert** event.|

## Remarks
<a name="sectionSection1"> </a>


 **Note**  Setting the value of a control by using a macro or Visual Basic doesn't trigger these events.

To run a macro or event procedure when these events occur, set the  **BeforeInsert** or **AfterInsert**property to the name of the macro or to [Event Procedure].

You can use an AfterInsert event procedure or macro to requery a recordset whenever a new record is added.

The BeforeInsert and AfterInsert events are similar to the  **BeforeUpdate**and  **AfterUpdate**events. These events occur in the following order:

 **BeforeInsert** â†’ **BeforeUpdate** â†’ **AfterUpdate** â†’ **AfterInsert**.

The following table summarizes the interaction between these events.



|**Event**|**Occurs when**|
|:-----|:-----|
|BeforeInsert|User types the first character in a new record.|
|BeforeUpdate|User updates the record.|
|AfterUpdate|Record is updated.|
|AfterInsert|Record updated is a new record.|
If the first character in a new record is typed into a text box or combo box, the  **BeforeInsert** event occurs before the **Change**event.


## Example
<a name="sectionSection2"> </a>

This example shows how you can use a  **BeforeInsert** event procedure to verify that the user wants to create a new record, and an **AfterInsert** event procedure to requery the record source for the Employees form after a record has been added.

To try the example, add the following event procedure to a form named Employees that is based on a table or query. Switch to form Datasheet view and try to insert a record.




```
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
<a name="sectionSection2"> </a>


#### Concepts


 [Form Object](72ef9219-142b-b690-b696-3eba9a5d4522.md)
#### Other resources


 [Form Object Members](e1976b58-28ca-8f76-cdf3-6732cb06ce6c.md)
