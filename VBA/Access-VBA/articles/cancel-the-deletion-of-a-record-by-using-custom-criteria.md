---
title: Cancel the Deletion of a Record by Using Custom Criteria
ms.prod: access
ms.assetid: 0445765f-4629-5970-776c-5bd30e2d72a1
ms.date: 06/08/2017
---


# Cancel the Deletion of a Record by Using Custom Criteria

The following example illutrates how to use a form's  **[Delete](form-delete-event-access.md)** event to prevent the deletion of a record based on custom criteria. In this example, the **Delete** event is canceled if the value of the DataRequired field is Yes.


```vb
Private Sub Form_Delete(Cancel As Integer) 
 
   ' Check the value of the DataRequired field. 
    If Me.DataRequired = "Yes" Then 
 
      ' Cancel the record deletion. 
      Cancel = True 
 
      ' Notify the user. 
       MsgBox "Cannot Delete the Record." 
    End If 
End Sub
```


