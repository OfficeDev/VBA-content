---
title: TextBox Control, AutoSize, Text Properties Example
keywords: fm20.chm5225181
f1_keywords:
- fm20.chm5225181
ms.prod: office
ms.assetid: a54a63a4-7428-2067-3eaa-aecf20d64428
ms.date: 06/08/2017
---


# TextBox Control, AutoSize, Text Properties Example

The following example demonstrates the effects of the  **AutoSize** property with a single-line **TextBox** and a multiline **TextBox**. The user can enter text into either **TextBox** and turn **AutoSize** on or off independently of the contents of the **TextBox**. This code sample also uses the **Text** property.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:




- Two  **TextBox** controls named TextBox1 and TextBox2.
    
- A  **ToggleButton** named ToggleButton1.
    




```vb
Private Sub UserForm_Initialize() 
 TextBox1.Text = "Single-line TextBox. " _ 
 &; "Type your text here." 
 
 TextBox2.MultiLine = True 
 TextBox2.Text = "Multi-line TextBox. Type " _ 
 &; "your text here. Use CTRL+ENTER to start " _ 
 &; "a new line." 
 
 ToggleButton1.Value = True 
 ToggleButton1.Caption = "AutoSize On" 
 TextBox1.AutoSize = True 
 TextBox2.AutoSize = True 
End Sub 
 
Private Sub ToggleButton1_Click() 
 If ToggleButton1.Value = True Then 
 ToggleButton1.Caption = "AutoSize On" 
 TextBox1.AutoSize = True 
 TextBox2.AutoSize = True 
 Else 
 ToggleButton1.Caption = "AutoSize Off" 
 TextBox1.AutoSize = False 
 TextBox2.AutoSize = False 
 End If 
End Sub
```


