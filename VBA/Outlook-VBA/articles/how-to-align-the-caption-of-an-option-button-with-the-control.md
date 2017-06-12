---
title: "How to: Align the Caption of an Option Button with the Control"
keywords: olfm10.chm3077154
f1_keywords:
- olfm10.chm3077154
ms.prod: outlook
ms.assetid: 4331a16a-6d73-855a-68d3-ef1fee6145bc
ms.date: 06/08/2017
---


# How to: Align the Caption of an Option Button with the Control

The following example demonstrates the  **[Alignment](optionbutton-alignment-property-outlook-forms-script.md)** property used with several **[OptionButton](optionbutton-object-outlook-forms-script.md)** controls. In this example, the user can change the alignment by clicking a **[ToggleButton](togglebutton-object-outlook-forms-script.md)**.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains the following controls:

- Two  **OptionButton** controls named OptionButton1 and OptionButton2.
    
- A  **ToggleButton** named ToggleButton1.
    



```vb
Dim OptionButton1 
Dim OptionButton2 
Dim ToggleButton1 
 
Sub Item_Open() 
 Set OptionButton1 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("OptionButton1") 
 Set OptionButton2 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("OptionButton2") 
 Set ToggleButton1 = Item.GetInspector.ModifiedFormPages.Item("P.2").Controls("ToggleButton1") 
 
 OptionButton1.Alignment = 0 'fmAlignmentLeft 
 OptionButton2.Alignment = 0 'fmAlignmentLeft 
 
 OptionButton1.Caption = "Alignment with AutoSize" 
 OptionButton2.Caption = "Choice 2" 
 OptionButton1.AutoSize = True 
 OptionButton2.AutoSize = True 
 
 ToggleButton1.Caption = "Left Align" 
 ToggleButton1.WordWrap = True 
 ToggleButton1.Value = True 
End Sub 
 
Sub ToggleButton1_Click() 
 If ToggleButton1.Value = True Then 
 ToggleButton1.Caption = "Left Align" 
 OptionButton1.Alignment = 0 'fmAlignmentLeft 
 OptionButton2.Alignment = 0 'fmAlignmentLeft 
 Else 
 ToggleButton1.Caption = "Right Align" 
 OptionButton1.Alignment = 1 'fmAlignmentRight 
 OptionButton2.Alignment = 1 'fmAlignmentRight 
 End If 
End Sub
```


