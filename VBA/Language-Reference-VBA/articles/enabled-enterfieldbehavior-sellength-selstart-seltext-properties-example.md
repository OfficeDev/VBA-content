---
title: Enabled, EnterFieldBehavior, SelLength, SelStart, SelText Properties Example
keywords: fm20.chm5225191
f1_keywords:
- fm20.chm5225191
ms.prod: office
ms.assetid: 3a21ec28-9d7e-1b11-9eb9-58907020ba79
ms.date: 06/08/2017
---


# Enabled, EnterFieldBehavior, SelLength, SelStart, SelText Properties Example

The following example tracks the selection-related properties ( **SelLength**, **SelStart**, and **SelText** ) that change as the user moves the insertion point and extends the selection using the keyboard. This example also uses the **Enabled** and **EnterFieldBehavior** properties.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:




- One large  **TextBox** named TextBox1.
    
- Three  **TextBox** controls in a column named TextBox2 through TextBox4.
    




```vb
Private Sub TextBox1_KeyUp(ByVal KeyCode As _ 
 MSForms.ReturnInteger, ByVal Shift As Integer) 
 TextBox2.Text = TextBox1.SelStart 
 TextBox3.Text = TextBox1.SelLength 
 TextBox4.Text = TextBox1.SelText 
End Sub
```




```vb
Private Sub UserForm_Initialize() 
 TextBox1.MultiLine = True 
 TextBox1.EnterFieldBehavior = _ 
 fmEnterFieldBehaviorRecallSelection 
 
 TextBox1.Text = "Type your text here. Use " _ 
 &; "CTRL+ENTER to start a new line." 
End Sub
```


