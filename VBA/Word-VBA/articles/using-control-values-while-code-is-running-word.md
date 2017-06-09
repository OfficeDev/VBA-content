---
title: Using Control Values While Code Is Running (Word)
keywords: vbawd10.chm5214008
f1_keywords:
- vbawd10.chm5214008
ms.prod: word
ms.assetid: 62722982-6725-57e2-099e-c31d0aefadd3
ms.date: 06/08/2017
---


# Using Control Values While Code Is Running (Word)

You can set and return some properties for  [ActiveX controls](http://msdn.microsoft.com/library/befa20c2-c4e7-1a53-7740-248885691710%28Office.15%29.aspx) while Visual Basic code is running. The following example sets the  **Text** property of a text box to "Hello."


```
TextBox1.Text = "Hello"
```


The data entered in a form by a user is lost when the form is closed. If you return the values of controls on a form after the form has been unloaded, you get the initial values for the controls rather than the values the user entered.

If you want to save the data entered in a form, you can save the information to module-level variables while the form is still running. The following example displays a form and saves the form data in public variables prior to unloading the form.



```vb
'Code in module to declare public variables 
Public strRegion As String 
Public intSalesPersonID As Integer 
Public blnCancelled As Boolean 
 
'Code in form 
Private Sub cmdCancel_Click() 
 Module1.blnCancelled = True 
 Unload Me 
End Sub 
 
Private Sub cmdOK_Click() 
 'Save data 
 intSalesPersonID = txtSalesPersonID.Text 
 strRegion = lstRegions.List(lstRegions.ListIndex) 
 Module1.blnCancelled = False 
 Unload Me 
End Sub 
 
Private Sub UserForm_Initialize() 
 Module1.blnCancelled = True 
End Sub 
 
'Code in module to display form 
Sub LaunchSalesPersonForm() 
 frmSalesPeople.Show 
 If blnCancelled = True Then 
 MsgBox "Operation Cancelled!", vbExclamation 
 Else 
 MsgBox "The Salesperson's ID is: " &; _ 
 intSalesPersonID &; _ 
 "The Region is: " &; strRegion 
 End If 
End Sub
```


