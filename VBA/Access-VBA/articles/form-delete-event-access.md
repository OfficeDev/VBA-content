---
title: Form.Delete Event (Access)
keywords: vbaac10.chm13639
f1_keywords:
- vbaac10.chm13639
ms.prod: access
api_name:
- Access.Form.Delete
ms.assetid: 89916f81-ec7a-f322-d4e6-a4a42db523cf
ms.date: 06/08/2017
---


# Form.Delete Event (Access)

Occurs when the user performs some action, such as pressing the DEL key, to delete a record, but before the record is actually deleted.


## Syntax

 _expression_. **Delete**( **_Cancel_**, )

 _expression_ A variable that represents a **Form** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required|**Integer**|The setting determines if the  **Delete** event occurs. Setting the Cancel argument to **True** (?1) cancels the **Delete** event.|

## Remarks

To run a macro or event procedure when these events occur, set the  **OnDelete**, **BeforeDelConfirm**, or **AfterDelConfirm** property to the name of the macro or to [Event Procedure].

After a record is deleted, it's stored in a temporary buffer. The  **BeforeDelConfirm** event occurs after the **Delete** event (or if you've deleted more than one record, after all the records are deleted, with a **Delete** event occurring for each record), but before the **Delete Confirm** dialog box is displayed. Canceling the **BeforeDelConfirm** event restores the record or records from the buffer and prevents the **Delete Confirm** dialog box from being displayed.

The  **AfterDelConfirm** event occurs after a record or records are actually deleted or after a deletion or deletions are canceled. If the **BeforeDelConfirm** event isn't canceled, the **AfterDelConfirm** event occurs after the **Delete Confirm** dialog box is displayed. The **AfterDelConfirm** event occurs even if the **BeforeDelConfirm** event is canceled. The **AfterDelConfirm** event procedure returns status information about the deletion. For example, you can use a macro or event procedure associated with the AfterDelConfirm event to recalculate totals affected by the deletion of records.

If you cancel the  **Delete** event, the **BeforeDelConfirm** and **AfterDelConfirm** events don't occur and the **Delete Confirm** dialog box isn't displayed.


 **Note**  The  **BeforeDelConfirm** and **AfterDelConfirm** events don't occur and the **Delete Confirm** dialog box isn't displayed if you clear the **Record Changes** check box under **Confirm** on the **Advanced** tab of the ** **Access Options**** dialog box, available by clicking the **Microsoft Office Button**
![File menu button](images/O12FileMenuButton_ZA10077102.gif)and the clicking  **Access Options**.

By running a macro or an event procedure when the  **Delete** event occurs, you can prevent a record from being deleted or allow a record to be deleted only under certain conditions. You can also use a **Delete** event to display a dialog box asking whether the user wants to delete a record before it's deleted.

To delete a record, you can click  **Delete Record** on the **Edit** menu. This deletes the current record (the record indicated by the record selector). You can also click the record selector or click **Select Record** on the **Edit** menu to select the record, and then press the DEL key to delete it. If you click **Delete Record**, the record selector of the current record, or **Select Record**, the **Exit** and **LostFocus** events for the control that has the focus occur. If you've changed any data in the record, the **BeforeUpdate** and **AfterUpdate** events for the record occur before the **Exit** and **LostFocus** events. If you click the record selector of a different record, the **Current** event for that record also occurs.

After you delete the record, the focus moves to the next record following the deleted record, and the Current event for that record occurs, followed by the  **Enter** and **GotFocus** events for the first control in that record.

The  **BeforeDelConfirm** event then occurs, just before Microsoft Access displays the **Delete Confirm** dialog box asking you to confirm the deletion. After you respond to the dialog box by confirming or canceling the deletion, the **AfterDelConfirm** event occurs.

You can delete one or more records at a time. The  **Delete** event occurs after each record is deleted. This enables you to access the data in each record before it's actually deleted, and selectively confirm or cancel each deletion in the **Delete** macro or event procedure. When you delete more than one record, the **Current** event for the record following the last deleted record and the **Enter** and **GotFocus** events for the first control in this record don't occur until all the records are deleted. In other words, a **Delete** event occurs for each selected record, but no other events occur until all the selected records are deleted. The **BeforeDelConfirm** and **AfterDelConfirm** events also don't occur until all the selected records are deleted.


## Example

The following example shows how you can prevent a user from deleting records from a table.

To try this example, add the following event procedure to a form that is based on a table. Switch to form Datasheet view and try to delete a record.




```vb
Private Sub Form_Delete(Cancel As Integer) 
    Cancel = True 
    MsgBox "This record can't be deleted." 
End Sub
```


## See also


#### Concepts


[Form Object](form-object-access.md)

