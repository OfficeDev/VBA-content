---
title: Locked, DropButtonStyle, ShowDropButtonWhen Properties Example
keywords: fm20.chm5225148
f1_keywords:
- fm20.chm5225148
ms.prod: office
ms.assetid: d661f20c-6bb1-e6e7-cbf9-bded76c549e6
ms.date: 06/08/2017
---


# Locked, DropButtonStyle, ShowDropButtonWhen Properties Example

The following example demonstrates the different symbols that you can specify for a drop-down arrow in a  **ComboBox** or **TextBox**. In this example, the user chooses a drop-down arrow style from a **ComboBox**. This example also uses the **Locked** property. To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:



- A  **ComboBox** named ComboBox1.
    
- A  **Label** named Label1.
    
- A  **TextBox** named TextBox1 placed beneath Label1.
    




```vb
Private Sub ComboBox1_Click() 
 ComboBox1.DropButtonStyle = ComboBox1.Value 
 TextBox1.DropButtonStyle = ComboBox1.Value 
End Sub 
 
Private Sub UserForm_Initialize() 
 ComboBox1.ColumnCount = 2 
 ComboBox1.BoundColumn = 2 
 ComboBox1.TextColumn = 1 
 
 ComboBox1.AddItem "Blank Button" 
 ComboBox1.List(0, 1) = 0 
 ComboBox1.AddItem "Down Arrow" 
 ComboBox1.List(1, 1) = 1 
 ComboBox1.AddItem "Ellipsis" 
 ComboBox1.List(2, 1) = 2 
 ComboBox1.AddItem "Underscore" 
 ComboBox1.List(3, 1) = 3 
 
 ComboBox1.Value = 0 
 
 TextBox1.Text = "TextBox1" 
 TextBox1.ShowDropButtonWhen = fmShowDropButtonWhenAlways 
 TextBox1.Locked = True 
 
 Label1.Caption = "TheDropButton also " _ 
 &; "applies to a TextBox." 
 Label1.AutoSize = True 
 Label1.WordWrap = False 
End Sub
```


