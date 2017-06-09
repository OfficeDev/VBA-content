---
title: "How to: Allow Multiple Lines, Word Wrapping, and Scrolling in a Text Box"
keywords: olfm10.chm3077220
f1_keywords:
- olfm10.chm3077220
ms.prod: outlook
ms.assetid: d478d32f-41b0-e64b-143a-3fac6f8b9624
ms.date: 06/08/2017
---


# How to: Allow Multiple Lines, Word Wrapping, and Scrolling in a Text Box

The following example demonstrates the  **[MultiLine](textbox-multiline-property-outlook-forms-script.md)**,  **[WordWrap](textbox-wordwrap-property-outlook-forms-script.md)**, and  **[ScrollBars](textbox-scrollbars-property-outlook-forms-script.md)** properties on a **[TextBox](textbox-object-outlook-forms-script.md)**.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:

- A  **TextBox** named TextBox1.
    
- Four  **[ToggleButton](togglebutton-object-outlook-forms-script.md)** controls named ToggleButton1 through ToggleButton4.
    
To see the entire text placed in the  **TextBox**, set  **MultiLine** and **WordWrap** to **True** by clicking the **ToggleButton** controls.
When  **MultiLine** is **True**, you can enter new lines of text by pressing SHIFT+ENTER.
 **ScrollBars** appears when you manually change the content of the **TextBox**.



```vb
Dim ToggleButton1 
Dim ToggleButton2 
Dim ToggleButton3 
Dim ToggleButton4 
Dim TextBox1 
 
Sub Item_Open 
'Initialize TextBox properties and toggle buttons 
 
 Set ToggleButton1 = Item.GetInspector.ModifiedFormPages("P.2").ToggleButton1 
 Set ToggleButton2 = Item.GetInspector.ModifiedFormPages("P.2").ToggleButton2 
 Set ToggleButton3 = Item.GetInspector.ModifiedFormPages("P.2").ToggleButton3 
 Set ToggleButton4 = Item.GetInspector.ModifiedFormPages("P.2").ToggleButton4 
 Set TextBox1 = Item.GetInspector.ModifiedFormPages("P.2").TextBox1 
 
 TextBox1.Text = "Type your text here. Enter SHIFT+ENTER to move to a new line." 
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
 
Sub ToggleButton1_Click 
'Set AutoSize property and associated ToggleButton 
 
 If ToggleButton1.Value = True Then 
 TextBox1.AutoSize = True 
 ToggleButton1.Caption = "AutoSize On" 
 Else 
 TextBox1.AutoSize = False 
 ToggleButton1.Caption = "AutoSize Off" 
 End if 
End Sub 
 
Sub ToggleButton2_Click 
'Set WordWrap property and associated ToggleButton 
 
 If ToggleButton2.Value = True Then 
 TextBox1.WordWrap = True 
 ToggleButton2.Caption = "WordWrap On" 
 Else 
 TextBox1.WordWrap = False 
 ToggleButton2.Caption = "WordWrap Off" 
 End if 
End Sub 
 
Sub ToggleButton3_Click 
'Set ScrollBars property and associated ToggleButton 
 
 If ToggleButton3.Value = True Then 
 TextBox1.ScrollBars = 3 
 ToggleButton3.Caption = "ScrollBars On" 
 Else 
 TextBox1.ScrollBars = 0 
 ToggleButton3.Caption = "ScrollBars Off" 
 End if 
End Sub 
 
Sub ToggleButton4_Click 
'Set MultiLine property and associated ToggleButton 
 
 If ToggleButton4.Value = True Then 
 TextBox1.MultiLine = True 
 ToggleButton4.Caption = "Multiple Lines" 
 Else 
 TextBox1.MultiLine = False 
 ToggleButton4.Caption = "Single Line" 
 End if 
 End Sub
```


