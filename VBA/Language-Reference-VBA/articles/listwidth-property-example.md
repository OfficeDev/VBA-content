---
title: ListWidth Property Example
keywords: fm20.chm5225141
f1_keywords:
- fm20.chm5225141
ms.prod: office
ms.assetid: c0247082-8767-be2a-9713-40942d0a0afd
ms.date: 06/08/2017
---


# ListWidth Property Example

The following example uses a  **SpinButton** to control the width of the drop-down list of a **ComboBox**. The user changes the value of the **SpinButton**, then clicks on the drop-down arrow of the **ComboBox** to display the list.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:




- A  **ComboBox** named ComboBox1.
    
- A  **SpinButton** named SpinButton1.
    
- A  **Label** named Label1.
    




```vb
Private Sub SpinButton1_Change() 
 ComboBox1.ListWidth = SpinButton1.Value 
 Label1.Caption = "ListWidth = " _ 
 &; SpinButton1.Value 
End Sub 
 
Private Sub UserForm_Initialize() 
 Dim i As Integer 
 
 For i = 1 To 20 
 ComboBox1.AddItem "Choice " _ 
 &; (ComboBox1.ListCount + 1) 
 Next i 
 
 SpinButton1.Min = 0 
 SpinButton1.Max = 130 
 SpinButton1.Value = Val(ComboBox1.ListWidth) 
 SpinButton1.SmallChange = 5 
 Label1.Caption = "ListWidth = " _ 
 &; SpinButton1.Value 
End Sub
```


