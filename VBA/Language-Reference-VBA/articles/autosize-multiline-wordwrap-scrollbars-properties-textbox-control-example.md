---
title: AutoSize, MultiLine, WordWrap, ScrollBars Properties, TextBox Control Example
keywords: fm20.chm5225173
f1_keywords:
- fm20.chm5225173
ms.prod: office
ms.assetid: aeac8985-2fe9-9fe8-6ad1-74e5322bc180
ms.date: 06/08/2017
---


# AutoSize, MultiLine, WordWrap, ScrollBars Properties, TextBox Control Example

The following example demonstrates the  **MultiLine**, **WordWrap**, and **ScrollBars** properties on a **TextBox**.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:




- A  **TextBox** named TextBox1.
    
- Four  **ToggleButton** controls named ToggleButton1 through ToggleButton4.
    

To see the entire text placed in the  **TextBox**, set **MultiLine** and **WordWrap** to **True** by clicking the **ToggleButton** controls.
When  **MultiLine** is **True**, you can enter new lines of text by pressing SHIFT+ENTER.
 **ScrollBars** appears when you manually change the content of the **TextBox**.



```vb
Private Sub UserForm_Initialize() 
'Initialize TextBox properties and toggle buttons 
 
 TextBox1.Text = "Type your text here. " 
 &; "Enter SHIFT+ENTER to move to a new line." 
 
 TextBox1.AutoSize = False 
 ToggleButton1.Caption = "AutoSize Off" 
 ToggleButton1.Value = False 
 ToggleButton1.AutoSize = True 
 
 TextBox1.WordWrap = False 
 ToggleButton2.Caption = "WordWrap Off" 
 ToggleButton2.Value = False 
 ToggleButton2.AutoSize = True 
 
 TextBox1.ScrollBars = 0 
 ToggleButton3.Caption = "ScrollBars Off" 
 ToggleButton3.Value = False 
 ToggleButton3.AutoSize = True 
 
 TextBox1.MultiLine = False 
 ToggleButton4.Caption = "Single Line" 
 ToggleButton4.Value = False 
 ToggleButton4.AutoSize = True 
 End Sub 
 
Private Sub ToggleButton1_Click() 
'Set AutoSize property and associated ToggleButton 
 
 If ToggleButton1.Value = True Then 
 TextBox1.AutoSize = True 
 ToggleButton1.Caption = "AutoSize On" 
 Else 
 TextBox1.AutoSize = False 
 ToggleButton1.Caption = "AutoSize Off" 
 End If 
End Sub
```




```vb
Private Sub ToggleButton2_Click() 
'Set WordWrap property and associated ToggleButton 
 
 If ToggleButton2.Value = True Then 
 TextBox1.WordWrap = True 
 ToggleButton2.Caption = "WordWrap On" 
 Else 
 TextBox1.WordWrap = False 
 ToggleButton2.Caption = "WordWrap Off" 
 End If 
End Sub
```




```vb
Private Sub ToggleButton3_Click() 
'Set ScrollBars property and associated ToggleButton 
 
 If ToggleButton3.Value = True Then 
 TextBox1.ScrollBars = 3 
 ToggleButton3.Caption = "ScrollBars On" 
 Else 
 TextBox1.ScrollBars = 0 
 ToggleButton3.Caption = "ScrollBars Off" 
 End If 
End Sub
```




```vb
Private Sub ToggleButton4_Click() 
'Set MultiLine property and associated ToggleButton 
 
 If ToggleButton4.Value = True Then 
 TextBox1.MultiLine = True 
 ToggleButton4.Caption = "Multiple Lines" 
 Else 
 TextBox1.MultiLine = False 
 ToggleButton4.Caption = "Single Line" 
 End If 
End Sub
```


