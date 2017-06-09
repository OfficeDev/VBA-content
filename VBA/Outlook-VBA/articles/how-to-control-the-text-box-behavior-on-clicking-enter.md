---
title: "How to: Control the Text Box' Behavior on Clicking Enter"
keywords: olfm10.chm3077185
f1_keywords:
- olfm10.chm3077185
ms.prod: outlook
ms.assetid: bc3329f9-b5f4-bbd9-19f1-8526342f406b
ms.date: 06/08/2017
---


# How to: Control the Text Box' Behavior on Clicking Enter

The following example uses the  **[EnterKeyBehavior](textbox-enterkeybehavior-property-outlook-forms-script.md)** property to control the effect of ENTER in a **[TextBox](textbox-object-outlook-forms-script.md)**. In this example, the user can specify either a single-line or multiline  **TextBox**.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:

- A  **TextBox** named TextBox1.
    
- Two  **[ToggleButton](togglebutton-object-outlook-forms-script.md)** controls named ToggleButton1 and ToggleButton2.
    



```vb
Dim TextBox1 
Dim ToggleButton1 
Dim ToggleButton2 
 
Sub Item_Open() 
 set TextBox1 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("TextBox1") 
 set ToggleButton1 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("ToggleButton1") 
 set ToggleButton2 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("ToggleButton2") 
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
 TextBox1.Text = "Type your text here. If EnterKeyBehavior is True,"&; _ 
 " press Enter to start a new line. Otherwise, press SHIFT+ENTER." 
End Sub 
 
Sub ToggleButton1_Click() 
 If ToggleButton1.Value = True Then 
 TextBox1.EnterKeyBehavior = True 
 ToggleButton1.Caption = "EnterKeyBehavior is True" 
 Else 
 TextBox1.EnterKeyBehavior = False 
 ToggleButton1.Caption = "EnterKeyBehavior is False" 
 End If 
End Sub 
 
Sub ToggleButton2_Click() 
 If ToggleButton2.Value = True Then 
 TextBox1.MultiLine = True 
 ToggleButton2.Caption = "MultiLine TextBox" 
 Else 
 TextBox1.MultiLine = False 
 ToggleButton2.Caption = "Single-line TextBox" 
 End If 
End Sub
```


