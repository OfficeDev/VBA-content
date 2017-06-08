---
title: Prompt a User Before Saving a Record
ms.prod: access
ms.assetid: 4b47967c-a043-cc8a-774f-1df0b529f29b
ms.date: 06/08/2017
---


# Prompt a User Before Saving a Record

The following example illustrates how to use the [BeforeUpdate](form-beforeupdate-event-access.md) event to prompt users to confirm their changes each time they save a record in a form.


```vb
Private Sub Form_BeforeUpdate(Cancel As Integer) 
   Dim strMsg As String 
   Dim iResponse As Integer 
 
   ' Specify the message to display. 
   strMsg = "Do you wish to save the changes?" &; Chr(10) 
   strMsg = strMsg &; "Click Yes to Save or No to Discard changes." 
 
   ' Display the message box. 
   iResponse = MsgBox(strMsg, vbQuestion + vbYesNo, "Save Record?") 
    
   ' Check the user's response. 
   If iResponse = vbNo Then 
    
      ' Undo the change. 
      DoCmd.RunCommand acCmdUndo 
 
      ' Cancel the update. 
      Cancel = True 
   End If 
End Sub
```


