---
title: Form.BeforeDelConfirm Event (Access)
keywords: vbaac10.chm13640
f1_keywords:
- vbaac10.chm13640
ms.prod: access
api_name:
- Access.Form.BeforeDelConfirm
ms.assetid: 36b9147a-6bfb-d386-117a-b65cc4659da8
ms.date: 06/08/2017
---


# Form.BeforeDelConfirm Event (Access)

The  **BeforeDelConfirm** event occurs after the user deletes to the buffer one or more records, but before Microsoft Access displays a dialog box asking the user to confirm the deletions.


## Syntax

 _expression_. **BeforeDelConfirm**( ** _Cancel_**, ** _Response_** )

 _expression_ A variable that represents a **Form** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Cancel_|Required|**Integer**|The setting determines if the  **BeforeDelConfirm** event occurs. Setting the _Cancel_ argument to **True** cancels the **BeforeDelConfirm** event and prevents the **Delete Confirm** dialog box from being displayed. If the event is canceled, the original records are restored, but the **AfterDelConfirm** event still occurs. If _Cancel_ is set to **True**, the _Response_ argument is ignored. If _Cancel_ is set to **False** (0), which it is by default, the value in the _Response_ argument is used by Microsoft Access to determine the type of response to the delete event.|
| _Response_|Required|**Integer**|An intrinsic constant that determines whether Microsoft Access displays the Delete Confirm dialog box asking if the record should be deleted.  **acDataErrContinue** continues without displaying the **Delete Confirm** dialog box. Setting the _Cancel_ argument to **False** and the _Response_ argument to **acDataErrContinue** enables Microsoft Access to delete records without prompting the user. **acDataErrDisplay** displays the **Delete Confirm** dialog box. The default value is **acDataErrDisplay**.|

## Remarks

To run a macro or event procedure when these events occur, set the  **BeforeDelConfirm** property to the name of the macro or to [Event Procedure].

After a record is deleted, it's stored in a temporary buffer. In a Microsoft Access database , the  **BeforeDelConfirm** event occurs after the **Delete** event (or if you've deleted more than one record, after all the records are deleted, with a **Delete** event occurring for each record), but before the **Delete Confirm** dialog box is displayed. Canceling the **BeforeDelConfirm** event restores the record or records from the buffer and prevents the **Delete Confirm** dialog box from being displayed.

In a Microsoft Access database, the  **AfterDelConfirm** event occurs after a record or records are actually deleted or after a deletion or deletions are canceled. If the **BeforeDelConfirm** event isn't canceled, the **AfterDelConfirm** event occurs after the **Delete Confirm** dialog box is displayed. The **AfterDelConfirm** event occurs even if the **BeforeDelConfirm** event is canceled.

If you cancel the  **Delete** event, the **BeforeDelConfirm** event does not occur and the **Delete Confirm** dialog box isn't displayed.

In a Microsoft Access project (.adp), the  **BeforeDelConfirm** event occurs before the **Delete** event. In order to avoid opening unnecessary transactions on Microsoft SQL Server, Access prompts you to confirm the deletion before opening the transaction. If you confirm the deletion, Access opens a transaction on Microsoft SQL Server, issues the DELETE statement to delete the record or records, and fires the form's **Delete** event. If you click **No** when prompted to confirm the deletion, Microsoft Access does not open a transaction on Microsoft SQL Server to delete the record and does not fire the form's **Delete** event.


 **Note**  The  **BeforeDelConfirm** event does not occur and the **Delete Confirm** dialog box isn't displayed if you clear the **Record Changes**check box under  **Confirm** on the **Editing** tab of the **Access Options** dialog box.

By running a macro or an event procedure when the  **Delete** event occurs, you can prevent a record from being deleted or allow a record to be deleted only under certain conditions. You can also use a **Delete** event to display a dialog box asking whether the user wants to delete a record before it's deleted.

After you delete the record, the focus moves to the next record following the deleted record, and the Current event for that record occurs, followed by the  **Enter** and **GotFocus** events for the first control in that record.

The  **BeforeDelConfirm** event then occurs, just before Microsoft Access displays the **Delete Confirm** dialog box asking you to confirm the deletion. After you respond to the dialog box by confirming or canceling the deletion, the **AfterDelConfirm** event occurs.

You can delete one or more records at a time. The  **Delete** event occurs after each record is deleted. This enables you to access the data in each record before it's actually deleted, and selectively confirm or cancel each deletion in the **Delete** macro or event procedure. When you delete more than one record, the **Current** event for the record following the last deleted record and the **Enter** and **GotFocus** events for the first control in this record don't occur until all the records are deleted. In other words, a **Delete** event occurs for each selected record, but no other events occur until all the selected records are deleted. The **BeforeDelConfirm** event does not occur until all the selected records are deleted.


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

