---
title: Form.AfterDelConfirm Event (Access)
keywords: vbaac10.chm13641
f1_keywords:
- vbaac10.chm13641
ms.prod: access
api_name:
- Access.Form.AfterDelConfirm
ms.assetid: 49f6f575-6f67-08b0-a2aa-913c8182cbe9
ms.date: 06/08/2017
---


# Form.AfterDelConfirm Event (Access)

The  **AfterDelConfirm** event occurs after the user confirms the deletions and the records are actually deleted or when the deletions are canceled.


## Syntax

 _expression_. **AfterDelConfirm**( ** _Status_** )

 _expression_ A variable that represents a **Form** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Status_|Required|**Integer**|An intrinsic constant that indicates whether a record has been deleted.  **acDeleteOK** indicates the deletion was successful. **acDeleteCancel** indicates the deletion was canceled in Visual Basic. **acDeleteUserCancel** indicates the deletion was canceled by the user.|

## Remarks

To run a macro or event procedure when these events occur, set the  **AfterDelConfirm** property to the name of the macro or to [Event Procedure].

After a record is deleted, it's stored in a temporary buffer.

In a Microsoft Access database, the  **AfterDelConfirm** event occurs after a record or records are actually deleted or after a deletion or deletions are canceled. If the **BeforeDelConfirm** event isn't canceled, the **AfterDelConfirm** event occurs after the **Delete Confirm** dialog box is displayed. The **AfterDelConfirm** event occurs even if the **BeforeDelConfirm** event is canceled. The **AfterDelConfirm** event procedure returns status information about the deletion. For example, you can use a macro or event procedure associated with the **AfterDelConfirm** event to recalculate totals affected by the deletion of records.

In a Microsoft Access project (.adp), the  **AfterDelConfirm** event occurs before a record or records are actually deleted. In order to avoid opening unnecessary transactions on Microsoft SQL Server, Access prompts you to confirm the deletion before opening the transaction. If you confirm the deletion, Access opens a transaction on Microsoft SQL Server, issues the DELETE statement to delete the record or records, and fires the form's Delete event. If you click **No** when prompted to confirm the deletion, Microsoft Access does not open a transaction on Microsoft SQL Server to delete the record and does not fire the form's Delete event.

If you cancel the  **Delete** event, the **AfterDelConfirm** event does not occur and the **Delete Confirm** dialog box isn't displayed.


 **Note**  The  **AfterDelConfirm** event does not occur and the **Delete Confirm** dialog box isn't displayed if you clear the **Record Changes**check box under  **Confirm** on the **Editing** tab of the **Access Options** dialog box.

By running a macro or an event procedure when the  **Delete** event occurs, you can prevent a record from being deleted or allow a record to be deleted only under certain conditions. You can also use a **Delete** event to display a dialog box asking whether the user wants to delete a record before it's deleted.

After you delete the record, the focus moves to the next record following the deleted record, and the Current event for that record occurs, followed by the  **Enter** and **GotFocus** events for the first control in that record.

The BeforeDelConfirm event then occurs, just before Microsoft Access displays the  **Delete Confirm** dialog box asking you to confirm the deletion. After you respond to the dialog box by confirming or canceling the deletion, the **AfterDelConfirm** event occurs.

You can delete one or more records at a time. The  **Delete** event occurs after each record is deleted. This enables you to access the data in each record before it's actually deleted, and selectively confirm or cancel each deletion in the **Delete** macro or event procedure. When you delete more than one record, the **Current** event for the record following the last deleted record and the **Enter** and **GotFocus** events for the first control in this record don't occur until all the records are deleted. In other words, a **Delete** event occurs for each selected record, but no other events occur until all the selected records are deleted. The **AfterDelConfirm** event also does not occur until all the selected records are deleted.


## Example

The following example shows how you can use the  **BeforeDelConfirm** event procedure to suppress the **Delete Confirm** dialog box and display a custom dialog box when a record is deleted. It also shows how you can use the **AfterDelConfirm** event procedure to display a message indicating whether the deletion progressed in the usual way or whether it was canceled in Visual Basic or by the user.


```vb
Private Sub Form_BeforeDelConfirm(Cancel As Integer, _ 
 Response As Integer) 
 ' Suppress default Delete Confirm dialog box. 
 Response = acDataErrContinue 
 ' Display custom dialog box. 
 If MsgBox("Delete this record?", vbOKCancel) = vbCancel Then 
 Cancel = True 
 End If 
End Sub 
 
Private Sub Form_AfterDelConfirm(Status As Integer) 
 Select Case Status 
 Case acDeleteOK 
 MsgBox "Deletion occurred normally." 
 Case acDeleteCancel 
 MsgBox "Programmer canceled the deletion." 
 Case acDeleteUserCancel 
 MsgBox "User canceled the deletion." 
 End Select 
End Sub
```


## See also


#### Concepts


[Form Object](form-object-access.md)

