---
title: Hide a Subform if the Main Form Contains No Records
ms.prod: access
ms.assetid: 20482340-0c86-71c9-3ba1-b9f515397fbc
ms.date: 06/08/2017
---


# Hide a Subform if the Main Form Contains No Records

The following example illustrates how to hide a subform named  _Orders_Subform_ if its main form does not contain any records. The code resides in the main form's **[Current](form-current-event-access.md)** event procedure.


```vb
Private Sub Form_Current() 
 
    With Me![Orders_Subform].Form 
     
        ' Check the RecordCount of the Subform. 
        If .RecordsetClone.RecordCount = 0 Then 
         
            ' Hide the subform. 
            .Visible = False 
         
        End If 
    End With 
End Sub
```


