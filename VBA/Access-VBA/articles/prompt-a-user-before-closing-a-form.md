---
title: Prompt a User Before Closing a Form
ms.prod: access
ms.assetid: 3a29f7c0-5692-49f0-bbfe-f9132d5b582f
ms.date: 06/08/2017
---


# Prompt a User Before Closing a Form

The following example illustrates how to prompt the user to verify that the form should closed.


```vb
Private Sub Form_Unload(Cancel As Integer) 
 If MsgBox("Are you sure that you want to close this form?", vbYesNo) = vbYes Then 
 Exit Sub 
 Else 
 Cancel = True 
 End If 
End Sub
```


