---
title: "How to: Control the Drop Button Style and When to Show the Drop Button"
keywords: olfm10.chm3077182
f1_keywords:
- olfm10.chm3077182
ms.prod: outlook
ms.assetid: 899b839e-f67e-1533-c0b6-28206e9af74a
ms.date: 06/08/2017
---


# How to: Control the Drop Button Style and When to Show the Drop Button

The following example demonstrates the different symbols that you can specify for a drop-down arrow in a  **[ComboBox](combobox-object-outlook-forms-script.md)** or **[TextBox](textbox-object-outlook-forms-script.md)**. In this example, the user chooses a drop-down arrow style from a  **ComboBox**. This example also uses the  ** [TextBox.Locked](olktextbox-locked-property-outlook.md)** property. To use this example, copy this sample code to the Script Editor of a form. Make sure that the form contains:


- A  **ComboBox** named ComboBox1.
    
- A  **[Label](label-object-outlook-forms-script.md)** named Label1.
    
- A  **TextBox** named TextBox1 placed beneath Label1.
    

```vb
Dim TextBox1 
Dim ComboBox1 
Dim Label1 
 
Sub ComboBox1_Click() 
 ComboBox1.DropButtonStyle = ComboBox1.Value 
 TextBox1.DropButtonStyle = ComboBox1.Value 
End Sub 
 
Sub Item_Open() 
 Set TextBox1 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("TextBox1") 
 Set ComboBox1 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("ComboBox1") 
 Set Label1 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("Label1") 
 
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
 TextBox1.ShowDropButtonWhen = 2 'fmShowDropButtonWhenAlways 
 TextBox1.Locked = True 
 
 Label1.Caption = "TheDropButton also applies to a TextBox." 
 Label1.AutoSize = True 
 Label1.WordWrap = False 
End Sub
```


