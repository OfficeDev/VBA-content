---
title: Initializing Control Properties
keywords: vbaxl10.chm5202086
f1_keywords:
- vbaxl10.chm5202086
ms.prod: excel
ms.assetid: 7d9d256c-c9e5-b45a-4da9-745d58cb666b
ms.date: 06/08/2017
---


# Initializing Control Properties

You can initialize  [controls](activex-controls.md)at run time by using Visual Basic code in a macro. For example, you could fill a list box, set text values, or set option buttons.

The following example uses the  **AddItem** method to add data to a list box. Then it sets the value of a text box and displays the form.



```vb
Private Sub GetUserName() 
 With UserForm1 
 .lstRegions.AddItem "North" 
 .lstRegions.AddItem "South" 
 .lstRegions.AddItem "East" 
 .lstRegions.AddItem "West" 
 .txtSalesPersonID.Text = "00000" 
 .Show 
 ' ... 
 End With 
End Sub
```

You can also use code in the  **Initialize** event of a form to set initial values for controls on the form. An advantage to setting initial control values in the **Initialize** event is that the initialization code stays with the form. You can copy the form to another project, and when you run the **Show** method to display the dialog box, the controls will be initialized.



```vb
Private Sub UserForm_Initialize() 
 UserForm1.lstNames.AddItem "Test One" 
 UserForm1.lstNames.AddItem "Test Two" 
 UserForm1.txtUserName.Text = "Default Name" 
End Sub
```


