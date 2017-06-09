---
title: ZOrder Method, TextBox Control Example
keywords: fm20.chm5225179
f1_keywords:
- fm20.chm5225179
ms.prod: office
ms.assetid: 54449312-f49f-20b9-05bb-6d8751d20e04
ms.date: 06/08/2017
---


# ZOrder Method, TextBox Control Example

The following example sets the z-order of a  **TextBox**, so the user can display the entire **TextBox** (by bringing it to the front of the z-order) or can place the **TextBox** behind other controls (by sending it to the back of the z-order).

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:




- Three  **TextBox** controls named TextBox1 through TextBox3.
    
- A  **ToggleButton** named ToggleButton1.
    




```vb
Private Sub ToggleButton1_Click() 
If ToggleButton1.Value = True Then 
 TextBox2.ZOrder(fmTop) 
 'Place TextBox2 on Top of z-order 
 
 'Update ToggleButton caption to identify next state 
 ToggleButton1.Caption = "Send TextBox2 to back" 
Else 
 TextBox2.ZOrder(1) 
 'Place TextBox2 on Bottom of z-order 
 
 'Update ToggleButton caption to identify next state 
 ToggleButton1.Caption = "Bring TextBox2 to front" 
End If 
End Sub
```




```vb
Private Sub UserForm_Initialize() 
'Set up text boxes to show z-order in the form 
TextBox1.Text = "TextBox 1" 
TextBox2.Text = "TextBox 2" 
TextBox3.Text = "TextBox 3" 
 
TextBox1.Height = 40 
TextBox2.Height = 40 
TextBox3.Height = 40 
 
TextBox1.Width = 60 
TextBox2.Width = 60 
TextBox3.Width = 60 
 
TextBox1.Left = 10 
TextBox1.Top = 10 
 
TextBox2.Left = 25 'Overlap TextBox2 on TextBox1 
TextBox2.Top = 25 
 
TextBox3.Left = 40 'Overlap TextBox3 on TextBox2, TextBox1 
TextBox3.Top = 40 
 
ToggleButton1.Value = False 
ToggleButton1.Caption = "Bring TextBox2 to Front" 
ToggleButton1.Left = 10 
ToggleButton1.Top = 90 
ToggleButton1.Width = 50 
ToggleButton1.Height = 50 
 
End Sub
```


