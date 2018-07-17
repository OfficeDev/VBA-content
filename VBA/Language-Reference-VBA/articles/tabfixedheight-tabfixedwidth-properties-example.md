---
title: TabFixedHeight, TabFixedWidth Properties Example
keywords: fm20.chm5225126
f1_keywords:
- fm20.chm5225126
ms.prod: office
ms.assetid: b856840c-1855-b871-33cc-e210489ce499
ms.date: 06/08/2017
---


# TabFixedHeight, TabFixedWidth Properties Example

The following example uses the  **TabFixedHeight** and **TabFixedWidth** properties to set the size of the tabs used in **MultiPage** and **TabStrip**. The user clicks the **SpinButton** controls to adjust the height and width of the tabs within the **MultiPage** and **TabStrip**.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:




- A  **MultiPage** named MultiPage1.
    
- A  **TabStrip** named TabStrip1.
    
- A  **Label** named Label1 for the width control.
    
- A  **SpinButton** named SpinButton1 for the width control.
    
- A  **TextBox** named TextBox1 for the width control.
    
- A  **Label** named Label2 for the height control.
    
- A  **SpinButton** named SpinButton2 for the height control.
    
- A  **TextBox** named TextBox2 for the height control.
    




```vb
Private Sub UpdateTabWidth() 
 TextBox1.Text = SpinButton1.Value 
 TabStrip1.TabFixedWidth = SpinButton1.Value 
 MultiPage1.TabFixedWidth = SpinButton1.Value 
End Sub 
 
Private Sub UpdateTabHeight() 
 TextBox2.Text = SpinButton2.Value 
 TabStrip1.TabFixedHeight = SpinButton2.Value 
 MultiPage1.TabFixedHeight = SpinButton2.Value 
End Sub 
 
Private Sub UserForm_Initialize() 
 MultiPage1.Style = fmTabStyleButtons 
 
 Label1.Caption = "Tab Width" 
 SpinButton1.Min = 0 
 SpinButton1.Max = _ 
 TabStrip1.Width / TabStrip1.Tabs.Count 
 SpinButton1.Value = 0 
 TextBox1.Locked = True 
 
 UpdateTabWidth 
 
 Label2.Caption = "Tab Height" 
 SpinButton2.Min = 0 
 SpinButton2.Max = TabStrip1.Height 
 SpinButton2.Value = 0 
 TextBox2.Locked = True 
 
 UpdateTabHeight 
End Sub 
 
Private Sub SpinButton1_SpinDown() 
 UpdateTabWidth 
End Sub 
 
Private Sub SpinButton1_SpinUp() 
 UpdateTabWidth 
End Sub 
 
Private Sub SpinButton2_SpinDown() 
 UpdateTabHeight 
End Sub 
 
Private Sub SpinButton2_SpinUp() 
 UpdateTabHeight 
End Sub
```


