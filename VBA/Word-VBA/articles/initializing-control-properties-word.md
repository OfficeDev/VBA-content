---
title: Initializing Control Properties (Word)
keywords: vbawd10.chm5211067
f1_keywords:
- vbawd10.chm5211067
ms.prod: word
ms.assetid: 18ae617c-6d51-ae79-be3c-1493ce4f6ef3
ms.date: 06/08/2017
---


# Initializing Control Properties (Word)

You can initialize  [ActiveX controls](http://msdn.microsoft.com/library/befa20c2-c4e7-1a53-7740-248885691710%28Office.15%29.aspx) at run time by using Visual Basic code in a macro. For example, you could fill a list box, set text values, or set option buttons.

The following example uses the Visual Basic  **AddItem** method to add data to a list box named lstRegions. Then it sets the value of a text box and displays the form.



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

You can also use code in the Visual Basic Initialize event of a form to set initial values for controls on the form. An advantage to setting initial control values in the Initialize event is that the initialization code stays with the form. You can copy the form to another project, and when you run the  **Show** method to display the dialog box, the controls will be initialized.



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


