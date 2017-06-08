---
title: CheckBox Control, Enabled, Locked Properties Example
keywords: fm20.chm5225146
f1_keywords:
- fm20.chm5225146
ms.prod: office
ms.assetid: 0733a3d8-4057-b308-4c25-0f5ef529b668
ms.date: 06/08/2017
---


# CheckBox Control, Enabled, Locked Properties Example

The following example demonstrates the  **Enabled** and **Locked** properties and how they complement each other. This example exposes each property independently with a **CheckBox**, so you observe the settings individually and combined. This example also includes a second **TextBox** so you can copy and paste information between the **TextBox** controls and verify the activities supported by the settings of these properties.


 **Note**  You can copy the selection to the Clipboard using CTRL+C and paste using CTRL+V.


To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:



- A  **TextBox** named TextBox1.
    
- Two  **CheckBox** controls named CheckBox1 and CheckBox2.
    
- A second  **TextBox** named TextBox2.
    




```vb
Private Sub CheckBox1_Change() 
 TextBox2.Text = "TextBox2" 
 TextBox1.Enabled = CheckBox1.Value 
End Sub 
 
Private Sub CheckBox2_Change() 
 TextBox2.Text = "TextBox2" 
 TextBox1.Locked = CheckBox2.Value 
End Sub 
 
Private Sub UserForm_Initialize() 
 TextBox1.Text = "TextBox1" 
 TextBox1.Enabled = True 
 TextBox1.Locked = False 
 
 CheckBox1.Caption = "Enabled" 
 CheckBox1.Value = True 
 
 CheckBox2.Caption = "Locked" 
 CheckBox2.Value = False 
 
 TextBox2.Text = "TextBox2" 
End Sub
```


