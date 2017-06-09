---
title: Layout Event, LayoutEffect Property, Move Method Example
keywords: fm20.chm5225128
f1_keywords:
- fm20.chm5225128
ms.prod: office
ms.assetid: c3585b29-d100-89a8-8e64-3afe5dbae8b2
ms.date: 06/08/2017
---


# Layout Event, LayoutEffect Property, Move Method Example

The following example moves a selected control on a form with the  **Move** method, and uses the **Layout** event and **LayoutEffect** property to identify the control that moved (and changed the layout of the **UserForm** ). The user clicks a control to move and then clicks the **CommandButton**. A message box displays the name of the control that is moving.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:




- A  **TextBox** named TextBox1.
    
- A  **ComboBox** named ComboBox1.
    
- An  **OptionButton** named OptionButton1.
    
- A  **CommandButton** named CommandButton1.
    
- A  **ToggleButton** named ToggleButton1.
    




```vb
Private Sub UserForm_Initialize() 
 CommandButton1.Caption = "Move current control" 
 CommandButton1.AutoSize = True 
 CommandButton1.TakeFocusOnClick = False 
 
 ToggleButton1.Caption = "Use Layout Event" 
 ToggleButton1.Value = True 
End Sub 
 
Private Sub CommandButton1_Click() 
 If ActiveControl.Name = "ToggleButton1" Then 
 'Keep it stationary 
 Else 
 'Move the control, using Layout event when 
 'ToggleButton1.Value is True 
 ActiveControl.Move 0, 0, , , _ 
 ToggleButton1.Value 
 End If 
End Sub 
 
Private Sub UserForm_Layout() 
 Dim MyControl As Control 
 
 MsgBox "In the Layout Event" 
 
 'Find the control that is moving. 
 For Each MyControl In Controls 
 If MyControl.LayoutEffect = _ 
 fmLayoutEffectInitiate Then 
 MsgBox MyControl.Name &; " is moving." 
 Exit For 
 End If 
 Next 
End Sub 
 
Private Sub ToggleButton1_Click() 
 If ToggleButton1.Value = True Then 
 ToggleButton1.Caption = "Use Layout Event" 
 Else 
 ToggleButton1.Caption = "No Layout Event" 
 End If 
End Sub
```


