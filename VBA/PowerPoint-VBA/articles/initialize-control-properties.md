---
title: Initialize Control Properties
keywords: vbapp10.chm5192842
f1_keywords:
- vbapp10.chm5192842
ms.prod: powerpoint
ms.assetid: d73b960d-bf78-1917-fc54-7b9b7cc7ca10
ms.date: 06/08/2017
---


# Initialize Control Properties

You can initialize controls at run time by using Visual Basic code in a macro. For example, you could fill a list box, set text values, or set option buttons.

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

You can also use code in the Initialize event of a form to set initial values for controls on the form. An advantage to setting initial control values in the Initialize event is that the initialization code stays with the form. You can copy the form to another project, and when you run the  **Show** method to display the dialog box, the controls will be initialized.



```vb
Private Sub UserForm_Initialize()
    With UserForm1
        With .lstRegions
           .AddItem "North"
            .AddItem "South"
            .AddItem "East"
            .AddItem "West"
        End With
        .txtSalesPersonID.Text = "00000"
    End With
End Sub
```


