---
title: Use Control Values While Code Is Running
keywords: vbapp10.chm5194031
f1_keywords:
- vbapp10.chm5194031
ms.prod: powerpoint
ms.assetid: a885309e-4525-c866-114f-994b56bf0488
ms.date: 06/08/2017
---


# Use Control Values While Code Is Running

Some control properties can be set and returned while Visual Basic code is running. The following example sets the  **Text** property of a text box to "Hello."


```
TextBox1.Text = "Hello"
```


The data entered on a form by a user is lost when the form is closed. If you return the values of controls on a form after the form has been unloaded, you get the initial values for the controls rather than the values the user entered.

If you want to save the data entered on a form, you can save the information to module-level variables while the form is still running. The following example displays a form and saves the form data.



```vb
'Code in module to declare public variables
Public strRegion As String
Public intSalesPersonID As Integer
Public blnCanceled As Boolean

'Code in form
Private Sub cmdCancel_Click()
    Module1.blnCanceled = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
    'Save data
    intSalesPersonID = txtSalesPersonID.Text
    strRegion = lstRegions.List(lstRegions.ListIndex)
    Module1.blnCanceled = False
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    Module1.blnCanceled = True
End Sub

'Code in module to display form
Sub LaunchSalesPersonForm()
    frmSalesPeople.Show
    If blnCanceled = True Then
        MsgBox "Operation Canceled!", vbExclamation
    Else
        MsgBox "The Salesperson's ID is: " &;
            intSalesPersonID &; _
            "The Region is: " &; strRegion
    End If
End Sub
```


