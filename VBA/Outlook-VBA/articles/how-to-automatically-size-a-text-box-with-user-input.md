---
title: "How to: Automatically Size a Text Box with User Input"
keywords: olfm10.chm3077157
f1_keywords:
- olfm10.chm3077157
ms.prod: outlook
ms.assetid: 573c8112-4c65-2411-afba-a7233baaa9aa
ms.date: 06/08/2017
---


# How to: Automatically Size a Text Box with User Input

The following example demonstrates the effects of the  **[AutoSize](textbox-autosize-property-outlook-forms-script.md)** property with a single-line **[TextBox](textbox-object-outlook-forms-script.md)** and a multiline **TextBox**. The user can enter text into either of the  **TextBox** controls and turn **AutoSize** on or off independently of the contents of the **TextBox**. This code sample also uses the  **[Text](textbox-text-property-outlook-forms-script.md)** property.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:

- Two  **TextBox** controls named TextBox1 and TextBox2.
    
- A  **[ToggleButton](togglebutton-object-outlook-forms-script.md)** named ToggleButton1.
    



```vb
Dim ToggleButton1 
Dim TextBox1 
Dim TextBox2 
 
Sub Item_Open() 
 Set ToggleButton1 = Item.GetInspector.ModifiedFormPages("P.2").ToggleButton1 
 Set TextBox1 = Item.GetInspector.ModifiedFormPages("P.2").TextBox1 
 Set TextBox2 = Item.GetInspector.ModifiedFormPages("P.2").TextBox2 
 
 TextBox1.Text = "Single-line TextBox. Type your text here." 
 
 TextBox2.MultiLine = True 
 TextBox2.Text = "Multi-line TextBox. Type your text here. Use SHIFT+ENTER to start a new line." 
 
 ToggleButton1.Value = True 
 ToggleButton1.Caption = "AutoSize On" 
 TextBox1.AutoSize = True 
 TextBox2.AutoSize = True 
End Sub 
 
Sub ToggleButton1_Click() 
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


