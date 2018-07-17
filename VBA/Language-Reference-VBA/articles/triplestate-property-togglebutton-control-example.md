---
title: TripleState Property, ToggleButton Control Example
keywords: fm20.chm5225131
f1_keywords:
- fm20.chm5225131
ms.prod: office
ms.assetid: f3f464c6-3bf2-2aae-ee6a-ead74c6b1289
ms.date: 06/08/2017
---


# TripleState Property, ToggleButton Control Example

The following example uses the  **TripleState** property to allow **Null** as a legal value of a **CheckBox** and a **ToggleButton**. The user controls the value of **TripleState** through ToggleButton2. The user can set the value of a **CheckBox** or **ToggleButton** based on the value of **TripleState**.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:




- A  **CheckBox** named CheckBox1.
    
- A  **ToggleButton** named ToggleButton1.
    
- A  **ToggleButton** named ToggleButton2.
    




```vb
Private Sub UserForm_Initialize() 
 CheckBox1.Caption = "Value is True" 
 CheckBox1.Value = True 
 CheckBox1.TripleState = False 
 
 ToggleButton1.Caption = "Value is True" 
 ToggleButton1.Value = True 
 ToggleButton1.TripleState = False 
 
 ToggleButton2.Value = False 
 ToggleButton2.Caption = "Triple State Off" 
End Sub 
 
Private Sub ToggleButton2_Click() 
 If ToggleButton2.Value = True Then 
 ToggleButton2.Caption = "Triple State On" 
 CheckBox1.TripleState = True 
 ToggleButton1.TripleState = True 
 Else 
 ToggleButton2.Caption = "Triple State Off" 
 CheckBox1.TripleState = False 
 ToggleButton1.TripleState = False 
 End If 
End Sub 
 
Private Sub CheckBox1_Change() 
 If IsNull(CheckBox1.Value) Then 
 CheckBox1.Caption = "Value is Null" 
 ElseIf CheckBox1.Value = False Then 
 CheckBox1.Caption = "Value is False" 
 ElseIf CheckBox1.Value = True Then 
 CheckBox1.Caption = "Value is True" 
 End If 
End Sub 
 
Private Sub ToggleButton1_Change() 
 If IsNull(ToggleButton1.Value) Then 
 ToggleButton1.Caption = "Value is Null" 
 ElseIf ToggleButton1.Value = False Then 
 ToggleButton1.Caption = "Value is False" 
 ElseIf ToggleButton1.Value = True Then 
 ToggleButton1.Caption = "Value is True" 
 End If 
End Sub
```


