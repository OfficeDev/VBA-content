---
title: Using Control Values While Code Is Running
keywords: vbaxl10.chm5205777
f1_keywords:
- vbaxl10.chm5205777
ms.prod: excel
ms.assetid: 71975020-fbda-69d4-42ad-eb6e7a3cb8f5
ms.date: 06/08/2017
---


# Using Control Values While Code Is Running

Some  **[controls](activex-controls.md)** properties can be set and returned while Visual Basic code is running. The following example sets the  **Text** property of a text box to "Hello."


```vb
TextBox1.Text = "Hello"
```


The data entered on a form by a user is lost when the form is closed. If you return the values of controls on a form after the form has been unloaded, you get the initial values for the controls rather than the values the user entered.

If you want to save the data entered on a form, you can save the information to module-level variables while the form is still running. The following example displays a form and saves the form data.



```vb
' Code in module to declare public variables. 
Public strRegion As String 
Public intSalesPersonID As Integer 
Public blnCancelled As Boolean 
 
' Code in form. 
Private Sub cmdCancel_Click() 
 Module1.blnCancelled = True 
 Unload Me 
End Sub 
 
Private Sub cmdOK_Click() 
 ' Save data. 
 intSalesPersonID = txtSalesPersonID.Text 
 strRegion = lstRegions.List(lstRegions.ListIndex) 
 Module1.blnCancelled = False 
 Unload Me 
End Sub 
 
Private Sub UserForm_Initialize() 
 Module1.blnCancelled = True 
End Sub 
 
' Code in module to display form. 
Sub LaunchSalesPersonForm() 
 frmSalesPeople.Show 
 If blnCancelled = True Then 
 MsgBox "Operation Cancelled!", vbExclamation 
 Else 
 MsgBox "The Salesperson's ID is: " &; 
 intSalesPersonID &; _ 
 "The Region is: " &; strRegion 
 End If 
End Sub
```


