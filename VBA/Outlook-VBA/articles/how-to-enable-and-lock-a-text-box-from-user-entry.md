---
title: "How to: Enable and Lock a Text Box from User Entry"
keywords: olfm10.chm3077184
f1_keywords:
- olfm10.chm3077184
ms.prod: outlook
ms.assetid: 354918d6-90f2-7e3f-cd72-2fa7681372ef
ms.date: 06/08/2017
---


# How to: Enable and Lock a Text Box from User Entry

The following example demonstrates the  **[Enabled](textbox-enabled-property-outlook-forms-script.md)** and **[Locked](textbox-locked-property-outlook-forms-script.md)** properties and how they complement each other. This example exposes each property independently with a **[CheckBox](checkbox-object-outlook-forms-script.md)**, so you observe the settings individually and combined. This example also includes a second  **[TextBox](textbox-object-outlook-forms-script.md)** so you can copy and paste information between the **TextBox** controls and verify the activities supported by the settings of these properties.


 **Note**  You can copy the selection to the Clipboard using CTRL+C and paste using CTRL+V.


To use this example, copy this sample code to the Script Editor of a form. Make sure that the form contains:


- A  **TextBox** named TextBox1.
    
- Two  **CheckBox** controls named CheckBox1 and CheckBox2.
    
- A second  **TextBox** named TextBox2.
    



```vb
Dim TextBox1 
Dim TextBox2 
Dim CheckBox1 
Dim CheckBox2 
 
Sub CheckBox1_Click() 
 TextBox2.Text = "TextBox2" 
 TextBox1.Enabled = CheckBox1.Value 
End Sub 
 
Sub CheckBox2_Click() 
 TextBox2.Text = "TextBox2" 
 TextBox1.Locked = CheckBox2.Value 
End Sub 
 
Sub Item_Open() 
 Set TextBox1 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("TextBox1") 
 Set TextBox2 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("TextBox2") 
 Set CheckBox1 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("CheckBox1") 
 Set CheckBox2 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("CheckBox2") 
 
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


