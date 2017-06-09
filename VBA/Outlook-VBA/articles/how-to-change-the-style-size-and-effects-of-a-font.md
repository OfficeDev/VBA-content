---
title: "How to: Change the Style, Size, and Effects of a Font"
keywords: olfm10.chm3077158
f1_keywords:
- olfm10.chm3077158
ms.prod: outlook
ms.assetid: 17225340-8da2-69b8-3255-d6c925f16aaf
ms.date: 06/08/2017
---


# How to: Change the Style, Size, and Effects of a Font

The following example demonstrates a  **[Font](font-object-outlook-forms-script.md)** object and the **[Bold](font-bold-property-outlook-forms-script.md)**,  **[Italic](font-italic-property-outlook-forms-script.md)**,  **[Size](font-size-property-outlook-forms-script.md)**,  **[Strikethrough](font-strikethrough-property-outlook-forms-script.md)**,  **[Underline](font-underline-property-outlook-forms-script.md)**, and  **[Weight](font-weight-property-outlook-forms-script.md)** properties related to fonts. You can manipulate font properties of an object directly or by using an alias, as this example also shows.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:

- A  **[Label](label-object-outlook-forms-script.md)** named Label1.
    
- Four  **[ToggleButton](togglebutton-object-outlook-forms-script.md)** controls named ToggleButton1 through ToggleButton4.
    
- A second  **Label** and a **[TextBox](textbox-object-outlook-forms-script.md)** named Label2 and TextBox1.
    



```vb
Dim MyFont 
Dim ToggleButton1 
Dim ToggleButton2 
Dim ToggleButton3 
Dim ToggleButton4 
Dim Label1 
Dim Label2 
Dim TextBox1 
 
Sub Item_Open() 
 Set MyPage = Item.GetInspector.ModifiedFormPages("P.2") 
 Set ToggleButton1 = MyPage.ToggleButton1 
 Set ToggleButton2 = MyPage.ToggleButton2 
 Set ToggleButton3 = MyPage.ToggleButton3 
 Set ToggleButton4 = MyPage.ToggleButton4 
 Set Label1 = MyPage.Label1 
 Set Label2 = MyPage.Label2 
 Set TextBox1 = MyPage.TextBox1 
 Set MyFont = Label1.Font 
 
 ToggleButton1.Value = True 
 ToggleButton1.Caption = "Bold On" 
 
 Label1.AutoSize = True 'Set size of Label1 
 Label1.AutoSize = False 
 
 ToggleButton2.Value = False 
 ToggleButton2.Caption = "Italic Off" 
 
 ToggleButton3.Value = False 
 ToggleButton3.Caption = "StrikeThrough Off" 
 
 ToggleButton4.Value = False 
 ToggleButton4.Caption = "Underline Off" 
 
 Label2.Caption = "Font Weight" 
 TextBox1.Text = Label1.Font.Weight 
 TextBox1.Enabled = False 
 
End Sub 
 
Sub ToggleButton1_Click() 
 If ToggleButton1.Value = True Then 
 MyFont.Bold = True 'Using MyFont alias to control font 
 ToggleButton1.Caption = "Bold On" 
 MyFont.Size = 22 'Increase the font size 
 Else 
 MyFont.Bold = False 
 ToggleButton1.Caption = "Bold Off" 
 MyFont.Size = 8 'Return font size to initial size 
 End If 
 
 TextBox1.Text = CStr(MyFont.Weight) 'Bold and Weight are related 
End Sub 
 
Sub ToggleButton2_Click() 
 If ToggleButton2.Value = True Then 
 Label1.Font.Italic = True 'Using Label1.Font directly 
 ToggleButton2.Caption = "Italic On" 
 Else 
 Label1.Font.Italic = False 
 ToggleButton2.Caption = "Italic Off" 
 End If 
End Sub 
 
Sub ToggleButton3_Click() 
 If ToggleButton3.Value = True Then 
 Label1.Font.Strikethrough = True 'Using Label1.Font directly 
 ToggleButton3.Caption = "StrikeThrough On" 
 Else 
 Label1.Font.Strikethrough = False 
 ToggleButton3.Caption = "StrikeThrough Off" 
 End If 
End Sub 
 
Sub ToggleButton4_Click() 
 If ToggleButton4.Value = True Then 
 MyFont.Underline = True 'Using MyFont alias for Label1.Font 
 ToggleButton4.Caption = "Underline On" 
 Else 
 Label1.Font.Underline = False 
 ToggleButton4.Caption = "Underline Off" 
 End If 
End Sub
```


