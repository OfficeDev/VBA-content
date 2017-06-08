---
title: KeyPress Event Example
keywords: fm20.chm5225133
f1_keywords:
- fm20.chm5225133
ms.prod: office
ms.assetid: a41bc663-9be3-fc6a-87b6-baa5c0a4a526
ms.date: 06/08/2017
---


# KeyPress Event Example

The following example uses the  **KeyPress** event to copy keystrokes from one **TextBox** to a second **TextBox**. The user types into the appropriately marked **TextBox**.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:




- Two  **TextBox** controls named TextBox1 and TextBox2.
    




```vb
Private Sub TextBox1_KeyPress(ByVal KeyAscii As _ 
 MSForms.ReturnInteger) 
 TextBox2.Text = TextBox2.Text &; Chr(KeyAscii) 
 
 'To handle keyboard combinations (using SHIFT, 
 'CONTROL, OPTION, COMMAND, and another key), 
 'or TAB or ENTER, use the KeyDown or KeyUp event. 
End Sub 
 
Private Sub UserForm_Initialize() 
 Move 0, 0, 570, 380 
 
 TextBox1.Move 30, 40, 220, 160 
 TextBox1.MultiLine = True 
 TextBox1.WordWrap = True 
 TextBox1.Text = "Type text here." 
 TextBox1.EnterKeyBehavior = True 
 
 
 TextBox2.Move 298, 40, 220, 160 
 TextBox2.MultiLine = True 
 TextBox2.WordWrap = True 
 TextBox2.Text = "Typed text copied here." 
 TextBox2.Locked = True 
 End Sub
```


