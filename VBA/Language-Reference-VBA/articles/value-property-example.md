---
title: Value Property Example
keywords: fm20.chm5225132
f1_keywords:
- fm20.chm5225132
ms.prod: office
ms.assetid: 7d98bbfa-9f19-b554-b327-554b12508b70
ms.date: 06/08/2017
---


# Value Property Example

The following example demonstrates the values that the different types of controls can have by displaying the  **Value** property of a selected control. The user chooses a control by pressing TAB or by clicking on the control. Depending on the type of control, the user can also specify a value for the control by typing in the text area of the control, by clicking one or more times on the control, or by selecting an item, page, or tab within the control. The user can display the value of the selected control by clicking the appropriately labeled **CommandButton**.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:




- A  **CommandButton** named CommandButton1.
    
- A  **TextBox** named TextBox1.
    
- A  **CheckBox** named CheckBox1.
    
- A  **ComboBox** named ComboBox1.
    
- A  **CommandButton** named CommandButton2.
    
- A  **ListBox** named ListBox1.
    
- A  **MultiPage** named MultiPage1.
    
- Two  **OptionButton** controls named OptionButton1 and OptionButton2.
    
- A  **ScrollBar** named ScrollBar1.
    
- A  **SpinButton** named SpinButton1.
    
- A  **TabStrip** named TabStrip1.
    
- A  **TextBox** named TextBox2.
    
- A  **ToggleButton** named ToggleButton1.
    




```vb
Dim i As Integer 
 
Private Sub CommandButton1_Click() 
 TextBox1.Text = "Value of " &; ActiveControl.Name _ 
 &; " is " &; ActiveControl.Value 
End Sub 
 
Private Sub UserForm_Initialize() 
 CommandButton1.Caption = "Get value of " _ 
 &; "current control" 
 CommandButton1.AutoSize = True 
 CommandButton1.TakeFocusOnClick = False 
 CommandButton1.TabStop = False 
 
 TextBox1.AutoSize = True 
 
 For i = 0 To 10 
 ComboBox1.AddItem "Choice " &; (i + 1) 
 ListBox1.AddItem "Selection " &; (100 - i) 
 Next i 
 
 CheckBox1.TripleState = True 
 ToggleButton1.TripleState = True 
 
 TextBox2.Text = "Enter text here." 
End Sub
```


