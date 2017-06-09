---
title: "How to: Move a Text Box in Front of or Behind Another Control"
keywords: olfm10.chm3077267
f1_keywords:
- olfm10.chm3077267
ms.prod: outlook
ms.assetid: 43e00491-39e4-5608-dc51-794be11ac721
ms.date: 06/08/2017
---


# How to: Move a Text Box in Front of or Behind Another Control

The following example sets the z-order of a  **[TextBox](textbox-object-outlook-forms-script.md)**, so the user can display the entire  **TextBox** (by bringing it to the front of the z-order) or can place the **TextBox** behind other controls (by sending it to the back of the z-order).

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:

- Three  **TextBox** controls named TextBox1 through TextBox3.
    
- A  **[ToggleButton](togglebutton-object-outlook-forms-script.md)** named ToggleButton1.
    



```vb
Dim ToggleButton1 
Dim TextBox1 
Dim TextBox2 
Dim TextBox3 
 
Sub Item_Open() 
Set TextBox1 = Item.GetInspector.ModifiedFormPages("P.2").TextBox1 
Set TextBox2 = Item.GetInspector.ModifiedFormPages("P.2").TextBox2 
Set TextBox3 = Item.GetInspector.ModifiedFormPages("P.2").TextBox3 
Set ToggleButton1 = Item.GetInspector.ModifiedFormPages("P.2").ToggleButton1 
 
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
 
Sub ToggleButton1_Click() 
If ToggleButton1.Value = True Then 
 TextBox2.ZOrder (fmTop) 'Place TextBox2 on Top of z-order 
 
 'Update ToggleButton caption to identify next state 
 ToggleButton1.Caption = "Send TextBox2 to back" 
Else 
 TextBox2.ZOrder (1) 'Place TextBox2 on Bottom of z-order 
 
 'Update ToggleButton caption to identify next state 
 ToggleButton1.Caption = "Bring TextBox2 to front" 
End If 
End Sub
```


