---
title: Display a Custom Dialog Box When the User Deletes a Record
ms.prod: access
ms.assetid: 512b324b-fe2f-b086-78d2-4c09933f5d25
ms.date: 06/08/2017
---


# Display a Custom Dialog Box When the User Deletes a Record

When you select a record on a form and delete it, Access displays a dialog box asking the user to confirm the deletion of the record. If you want, you can prevent this dialog box from appearing in two ways. You can cancel the [BeforeDelConfirm](form-beforedelconfirm-event-access.md) event, in which case the deletion is canceled. Or you can set the _Response_ argument of the **BeforeDelConfirm** event procedure to **acDataErrContinue**, in which case the deletion is confirmed.

You can use a  **BeforeDelConfirm** event procedure to display a custom dialog box and handle users' responses. The following example demonstrates how to use a custom dialog box to ask users whether they want to cancel or proceed with the record deletion.



```vb
Private Sub Form_BeforeDelConfirm(Cancel As Integer, Response As Integer) 
 
   Dim strMessage As String 
   Dim intResponse As Integer 
 
On Error GoTo ErrorHandler 
 
   ' Display the custom dialog box. 
   strMessage = "Would you like to delete the current record?" 
   intResponse = MsgBox(strMessage, vbYesNo + vbQuestion, _ 
               "Continue delete?") 
 
   ' Check the response. 
   If intResponse = vbYes Then 
      Response = acDataErrContinue 
   Else 
      Cancel = True 
   End If 
 
   Exit Sub 
 
ErrorHandler: 
   MsgBox "Error #: " &; Err.Number &; vbCrLf &; vbCrLf &; Err.Description 
End Sub
```


