---
title: EnterKeyBehavior, MultiLine Property Example
keywords: fm20.chm5225144
f1_keywords:
- fm20.chm5225144
ms.prod: office
ms.assetid: 06f7eb5f-cb91-6231-ccf5-1dcdf57fb3c1
ms.date: 06/08/2017
---


# EnterKeyBehavior, MultiLine Property Example

The following example uses the  **EnterKeyBehavior** property to control the effect of ENTER in a **TextBox**. In this example, the user can specify either a single-line or multiline **TextBox**.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:




- A  **TextBox** named TextBox1.
    
- Two  **ToggleButton** controls named ToggleButton1 and ToggleButton2.
    




```vb
Private Sub UserForm_Initialize() 
 TextBox1.EnterKeyBehavior = True 
 ToggleButton1.Caption = "EnterKeyBehavior is True" 
 ToggleButton1.Width = 70 
 ToggleButton1.Value = True 
 
 TextBox1.MultiLine = True 
 ToggleButton2.Caption = "MultiLine is True" 
 ToggleButton2.Width = 70 
 ToggleButton2.Value = True 
 
 TextBox1.Height = 100 
 TextBox1.WordWrap = True 
 TextBox1.Text = "Type your text here. If " _ 
 &; "EnterKeyBehavior is True, " _ 
 &; "press Enter to start a new line. Otherwise, press SHIFT+ENTER." 
End Sub 
 
Private Sub ToggleButton1_Click() 
 If ToggleButton1.Value = True Then 
 TextBox1.EnterKeyBehavior = True 
 ToggleButton1.Caption = _ 
 "EnterKeyBehavior is True" 
 Else 
 TextBox1.EnterKeyBehavior = False 
 ToggleButton1.Caption = _ 
 "EnterKeyBehavior is False" 
 End If 
End Sub 
 
Private Sub ToggleButton2_Click() 
 If ToggleButton2.Value = True Then 
 TextBox1.MultiLine = True 
 ToggleButton2.Caption = "MultiLine TextBox" 
 Else 
 TextBox1.MultiLine = False 
 ToggleButton2.Caption = "Single-line TextBox" 
 End If 
End Sub
```


