---
title: DragBehavior, EnterFieldBehavior Properties Example
keywords: fm20.chm5225165
f1_keywords:
- fm20.chm5225165
ms.prod: office
ms.assetid: 3a422742-87c7-6d8d-493d-52942c383328
ms.date: 06/08/2017
---


# DragBehavior, EnterFieldBehavior Properties Example

The following example uses the  **DragBehavior** and **EnterFieldBehavior** properties to demonstrate the different effects that you can provide when entering a control and when dragging information from one control to another.

The sample uses two  **TextBox** controls. You can set **DragBehavior** and **EnterFieldBehavior** for each control and see the effects of dragging from one control to another.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:



- A  **TextBox** named TextBox1.
    
- Two  **ToggleButton** controls named ToggleButton1 and ToggleButton2. These controls are associated with TextBox1.
    
- A  **TextBox** named TextBox2.
    
- Two  **ToggleButton** controls named ToggleButton3 and ToggleButton4. These controls are associated with TextBox2.
    




```vb
Private Sub UserForm_Initialize() 
 TextBox1.Text = "Once upon a time in a land ...," 
 ToggleButton1.Value = True 
 ToggleButton1.Caption = "Drag Enabled" 
 ToggleButton1.WordWrap = True 
 TextBox1.DragBehavior = fmDragBehaviorEnabled 
 
 ToggleButton2.Value = True 
 ToggleButton2.Caption = "Recall Selection" 
 ToggleButton2.WordWrap = True 
 TextBox1.EnterFieldBehavior = _ 
 fmEnterFieldBehaviorRecallSelection 
 
 TextBox2.Text = "XXX, YYYY" 
 ToggleButton3.Value = False 
 ToggleButton3.Caption = "Drag Disabled" 
 ToggleButton3.WordWrap = True 
 TextBox2.DragBehavior = fmDragBehaviorDisabled 
 
 ToggleButton4.Value = False 
 ToggleButton4.Caption = "Select All" 
 ToggleButton4.WordWrap = True 
 TextBox2.EnterFieldBehavior = _ 
 fmEnterFieldBehaviorSelectAll 
End Sub 
 
Private Sub ToggleButton1_Click() 
 If ToggleButton1.Value = True Then 
 ToggleButton1.Caption = "Drag Enabled" 
 TextBox1.DragBehavior = fmDragBehaviorEnabled 
 Else 
 ToggleButton1.Caption = "Drag Disabled" 
 TextBox1.DragBehavior = fmDragBehaviorDisabled 
 End If 
End Sub 
 
Private Sub ToggleButton2_Click() 
 If ToggleButton2.Value = True Then 
 ToggleButton2.Caption = "Recall Selection" 
 TextBox1.EnterFieldBehavior = _ 
 fmEnterFieldBehaviorRecallSelection 
 Else 
 ToggleButton2.Caption = "Select All" 
 TextBox1.EnterFieldBehavior = _ 
 fmEnterFieldBehaviorSelectAll 
 End If 
End Sub 
 
Private Sub ToggleButton3_Click() 
 If ToggleButton3.Value = True Then 
 ToggleButton3.Caption = "Drag Enabled" 
 TextBox2.DragBehavior = fmDragBehaviorEnabled 
 Else 
 ToggleButton3.Caption = "Drag Disabled" 
 TextBox2.DragBehavior = fmDragBehaviorDisabled 
 End If 
End Sub 
 
Private Sub ToggleButton4_Click() 
 If ToggleButton4.Value = True Then 
 ToggleButton4.Caption = "Recall Selection" 
 TextBox2.EnterFieldBehavior = _ 
 fmEnterFieldBehaviorRecallSelection 
 Else 
 ToggleButton4.Caption = "Select All" 
 TextBox2.EnterFieldBehavior = _ 
 fmEnterFieldBehaviorSelectAll 
 End If 
End Sub
```


