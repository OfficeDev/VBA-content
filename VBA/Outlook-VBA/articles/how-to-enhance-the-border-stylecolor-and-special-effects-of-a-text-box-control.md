---
title: "How to: Enhance the Border Style, Color, and Special Effects of a Text Box Control"
keywords: olfm10.chm3077159
f1_keywords:
- olfm10.chm3077159
ms.prod: outlook
ms.assetid: 250de388-e1e8-98a6-95bd-df3ff3eb6a0a
ms.date: 06/08/2017
---


# How to: Enhance the Border Style, Color, and Special Effects of a Text Box Control

The following example demonstrates the  **[BorderColor](textbox-bordercolor-property-outlook-forms-script.md)** and **[SpecialEffect](textbox-specialeffect-property-outlook-forms-script.md)** properties, showing each border available through these properties. The example also demonstrates how to control color settings by using the **[BackColor](textbox-backcolor-property-outlook-forms-script.md)**,  **[BackStyle](textbox-backstyle-property-outlook-forms-script.md)**,  **BorderColor**, and  **[ForeColor](textbox-forecolor-property-outlook-forms-script.md)** properties.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:

- Six  **[TextBox](textbox-object-outlook-forms-script.md)** controls named TextBox1 through TextBox6.
    
- Two  **[ToggleButton](togglebutton-object-outlook-forms-script.md)** controls named ToggleButton1 and ToggleButton2.
    



```vb
Dim TextBox1 
Dim TextBox2 
Dim TextBox3 
Dim TextBox4 
Dim TextBox5 
Dim TextBox6 
Dim ToggleButton1 
Dim ToggleButton2 
 
Sub Item_Open() 
Set TextBox1 = Item.GetInspector.ModifiedFormPages("P.2").TextBox1 
Set TextBox2 = Item.GetInspector.ModifiedFormPages("P.2").TextBox2 
Set TextBox3 = Item.GetInspector.ModifiedFormPages("P.2").TextBox3 
Set TextBox4 = Item.GetInspector.ModifiedFormPages("P.2").TextBox4 
Set TextBox5 = Item.GetInspector.ModifiedFormPages("P.2").TextBox5 
Set TextBox6 = Item.GetInspector.ModifiedFormPages("P.2").TextBox6 
Set ToggleButton1 = Item.GetInspector.ModifiedFormPages("P.2").ToggleButton1 
Set ToggleButton2 = Item.GetInspector.ModifiedFormPages("P.2").ToggleButton2 
'Initialize each TextBox with a border style or special effect, 
'and foreground and background colors 
 
'TextBox1 initially uses a borderstyle 
TextBox1.Text = "BorderStyle-Single" 
TextBox1.BorderStyle = 1 
TextBox1.BorderColor = RGB(255, 128, 128) 'Color - Salmon 
TextBox1.ForeColor = RGB(255, 255, 0) 'Color - Yellow 
TextBox1.BackColor = RGB(0, 128, 64) 'Color - Green #2 
 
'TextBoxes 2 through 6 initially use special effects 
TextBox2.Text = "Flat" 
TextBox2.SpecialEffect = 0 
TextBox2.ForeColor = RGB(64, 0, 0) 'Color - Brown 
TextBox2.BackColor = RGB(0, 0, 255) 'Color - Blue 
 
'Ensure the background style for TextBox2 is initially opaque. 
TextBox2.BackStyle = 1 
 
TextBox3.Text = "Etched" 
TextBox3.SpecialEffect = 3 
TextBox3.ForeColor = RGB(128, 0, 255) 'Color - Purple 
TextBox3.BackColor = RGB(0, 255, 255) 'Color - Cyan 
 
'Define BorderColor for later use (when borderstyle=fmBorderStyleSingle) 
TextBox3.BorderColor = RGB(0, 0, 0) 'Color - Black 
 
TextBox4.Text = "Bump" 
TextBox4.SpecialEffect = 6 
TextBox4.ForeColor = RGB(255, 0, 255) 'Color - Magenta 
TextBox4.BackColor = RGB(0, 0, 100) 'Color - Navy blue 
 
TextBox5.Text = "Raised" 
TextBox5.SpecialEffect = 1 
TextBox5.ForeColor = RGB(255, 0, 0) 'Color - Red 
TextBox5.BackColor = RGB(128, 128, 128) 'Color - Gray 
 
TextBox6.Text = "Sunken" 
TextBox6.SpecialEffect = 2 
TextBox6.ForeColor = RGB(0, 64, 0) 'Color - Olive 
TextBox6.BackColor = RGB(0, 255, 0) 'Color - Green #1 
 
ToggleButton1.Caption = "Swap styles" 
ToggleButton2.Caption = "Transparent/Opaque background" 
End Sub 
 
Sub ToggleButton1_Click() 
 
'Swap borders between TextBox1 and TextBox3 
If ToggleButton1.Value = True Then 
 'Change TextBox1 from BorderStyle to Etched 
 TextBox1.Text = "Etched" 
 TextBox1.SpecialEffect = 3 
 
 'Change TextBox3 from Etched to BorderStyle 
 TextBox3.Text = "BorderStyle-Single" 
 TextBox3.BorderStyle = 1 
Else 
 'Change TextBox1 back to BorderStyle 
 TextBox1.Text = "BorderStyle-Single" 
 TextBox1.BorderStyle = 1 
 
 'Change TextBox3 back to Etched 
 TextBox3.Text = "Etched" 
 TextBox3.SpecialEffect = 3 
End If 
End Sub 
 
 
Sub ToggleButton2_Click() 
 
'Set background to Opaque or Transparent 
If ToggleButton2.Value = True Then 
 'Change TextBox2 to a transparent background 
 TextBox2.BackStyle = 0 
Else 
 'Change TextBox2 back to opaque background 
 TextBox2.BackStyle = 1 
End If 
 
End Sub
```


