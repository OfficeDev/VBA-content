---
title: Style Property Example
keywords: fm20.chm5225123
f1_keywords:
- fm20.chm5225123
ms.prod: office
ms.assetid: ca09e7da-1b5f-f106-96c8-ec1a7b4ef2a0
ms.date: 06/08/2017
---


# Style Property Example

The following example uses the  **Style** property to change the effect of typing in the text area of a **ComboBox**. The user chooses a style by selecting an **OptionButton** control and then types into the **ComboBox** to select an item. When **Style** is _fmStyleDropDownList_, the user must choose an item from the drop-down list. When **Style** is _fmStyleDropDownCombo_, the user can type into the text area to specify an item in the drop-down list.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:




- Two  **OptionButton** controls named OptionButton1 and OptionButton2.
    
- A  **ComboBox** named ComboBox1.
    




```vb
Private Sub OptionButton1_Click() 
 ComboBox1.Style = fmStyleDropDownCombo 
End Sub 
 
Private Sub OptionButton2_Click() 
 ComboBox1.Style = fmStyleDropDownList 
End Sub 
 
Private Sub UserForm_Initialize() 
 Dim i As Integer 
 
 For i = 1 To 10 
 ComboBox1.AddItem "Choice " &; i 
 Next i 
 
 OptionButton1.Caption = "Select like ComboBox" 
 OptionButton1.Value = True 
 ComboBox1.Style = fmStyleDropDownCombo 
 
 OptionButton2.Caption = "Select like ListBox" 
End Sub
```


