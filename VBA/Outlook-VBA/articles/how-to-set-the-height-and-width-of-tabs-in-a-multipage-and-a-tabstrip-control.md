---
title: "How to: Set the Height and Width of Tabs in a MultiPage and a TabStrip Control"
keywords: olfm10.chm3077249
f1_keywords:
- olfm10.chm3077249
ms.prod: outlook
ms.assetid: 4d351a3d-334e-5356-7a2d-6c7b11655319
ms.date: 06/08/2017
---


# How to: Set the Height and Width of Tabs in a MultiPage and a TabStrip Control

The following example uses the  **TabFixedHeight** and **TabFixedWidth** properties to set the size of the tabs used in **[MultiPage](multipage-object-outlook-forms-script.md)** and **[TabStrip](tabstrip-object-outlook-forms-script.md)**. The user clicks the  **[SpinButton](spinbutton-object-outlook-forms-script.md)** controls to adjust the height and width of the tabs within the **MultiPage** and **TabStrip**.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:

- A  **MultiPage** named MultiPage1.
    
- A  **TabStrip** named TabStrip1.
    
- A  **[Label](label-object-outlook-forms-script.md)** named Label1 for the width control.
    
- A  **SpinButton** named SpinButton1 for the width control that is bound to a custom number field named SpinButtonWidth.
    
- A  **[TextBox](textbox-object-outlook-forms-script.md)** named TextBox1 for the width control.
    
- A  **Label** named Label2 for the height control.
    
- A  **SpinButton** named SpinButton2 for the height control that is bound to a custom number field named SpinButtonHeight.
    
- A  **TextBox** named TextBox2 for the height control.
    



```vb
Sub UpdateTabWidth() 
 Set TextBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TextBox1") 
 Set SpinButton1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("SpinButton1") 
 Set TabStrip1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TabStrip1") 
 Set MultiPage1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("MultiPage1") 
 
 TextBox1.Text = SpinButton1.Value 
 TabStrip1.TabFixedWidth = SpinButton1.Value 
 MultiPage1.TabFixedWidth = SpinButton1.Value 
End Sub 
 
Sub UpdateTabHeight() 
 Set TextBox2 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TextBox2") 
 Set SpinButton2 = Item.GetInspector.ModifiedFormPages("P.2").Controls("SpinButton2") 
 Set TabStrip1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TabStrip1") 
 Set MultiPage1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("MultiPage1") 
 
 TextBox2.Text = SpinButton2.Value 
 TabStrip1.TabFixedHeight = SpinButton2.Value 
 MultiPage1.TabFixedHeight = SpinButton2.Value 
End Sub 
 
Sub Item_Open() 
 Set TextBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TextBox1") 
 Set TextBox2 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TextBox2") 
 Set SpinButton1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("SpinButton1") 
 Set SpinButton2 = Item.GetInspector.ModifiedFormPages("P.2").Controls("SpinButton2") 
 Set TabStrip1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TabStrip1") 
 Set MultiPage1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("MultiPage1") 
 Set Label1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("Label1") 
 Set Label2 = Item.GetInspector.ModifiedFormPages("P.2").Controls("Label2") 
 
 MultiPage1.Style = 1 '1=fmTabStyleButtons 
 
 Label1.Caption = "Tab Width" 
 SpinButton1.Min = 0 
 SpinButton1.Max = TabStrip1.Width / TabStrip1.Tabs.Count 
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
 
Sub Item_CustomPropertyChange(byval pname) 
'msgbox pname 
 If pname = "SpinButtonWidth" Then 
 UpdateTabWidth 
 ElseIf pname = "SpinButtonHeight" Then 
 UpdateTabHeight 
 End If 
End Sub
```


