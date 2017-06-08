---
title: Alignment Property Example
keywords: fm20.chm5225151
f1_keywords:
- fm20.chm5225151
ms.prod: office
ms.assetid: cd079dcf-c5d2-e259-3607-a2a8e2864e02
ms.date: 06/08/2017
---


# Alignment Property Example

The following example demonstrates the  **Alignment** property used with several **OptionButton** controls. In this example, the user can change the alignment by clicking a **ToggleButton**.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains the following controls:




- Two  **OptionButton** controls named OptionButton1 and OptionButton2.
    
- A  **ToggleButton** named ToggleButton1.
    




```vb
Private Sub UserForm_Initialize() 
 OptionButton1.Alignment = fmAlignmentLeft 
 OptionButton2.Alignment = fmAlignmentLeft 
 
 OptionButton1.Caption = "Alignment with AutoSize" 
 OptionButton2.Caption = "Choice 2" 
 OptionButton1.AutoSize = True 
 OptionButton2.AutoSize = True 
 
 ToggleButton1.Caption = "Left Align" 
 ToggleButton1.WordWrap = True 
 ToggleButton1.Value = True 
End Sub 
 
Private Sub ToggleButton1_Click() 
 If ToggleButton1.Value = True Then 
 ToggleButton1.Caption = "Left Align" 
 OptionButton1.Alignment = fmAlignmentLeft 
 OptionButton2.Alignment = fmAlignmentLeft 
 Else 
 ToggleButton1.Caption = "Right Align" 
 OptionButton1.Alignment = fmAlignmentRight 
 OptionButton2.Alignment = fmAlignmentRight 
 End If 
End Sub
```


