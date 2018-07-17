---
title: KeyDown Event, TextBox Control, CurLine, CurTargetX, CurX, Text Properties Example
keywords: fm20.chm5225187
f1_keywords:
- fm20.chm5225187
ms.prod: office
ms.assetid: 696c6429-7a62-9eeb-d7c3-a883e888da09
ms.date: 06/08/2017
---


# KeyDown Event, TextBox Control, CurLine, CurTargetX, CurX, Text Properties Example

The following example tracks the  **CurLine**, **CurTargetX**, and **CurX** property settings in a multiline **TextBox**. These settings change in the **KeyUp** event as the user types into the **Text** property, moves the insertion point, and extends the selection using the keyboard.

To use this example, follow these steps:




1. Copy this sample code to the Declarations portion of a form.
    
2. Add one large  **TextBox** named TextBox1 to the form.
    
3. Add three  **TextBox** controls named TextBox2, TextBox3, and TextBox4 in a column.
    




```vb
Private Sub TextBox1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer) 
 TextBox2.Text = TextBox1.CurLine 
 TextBox3.Text = TextBox1.CurX 
 TextBox4.Text = TextBox1.CurTargetX 
End Sub
```




```vb
Private Sub UserForm_Initialize() 
 TextBox1.MultiLine = True 
 
 TextBox1.Text = "Type your text here. User CTRL + ENTER to start a new line." 
End Sub
```


