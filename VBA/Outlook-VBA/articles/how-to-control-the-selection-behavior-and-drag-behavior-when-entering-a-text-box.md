---
title: "How to: Control the Selection Behavior and Drag Behavior When Entering a Text Box"
keywords: olfm10.chm3077181
f1_keywords:
- olfm10.chm3077181
ms.prod: outlook
ms.assetid: 81d54db0-0bfe-3e21-b3ea-643980c8f48b
ms.date: 06/08/2017
---


# How to: Control the Selection Behavior and Drag Behavior When Entering a Text Box

The following example uses the  **[DragBehavior](textbox-dragbehavior-property-outlook-forms-script.md)** and **[EnterFieldBehavior](olktextbox-enterfieldbehavior-property-outlook.md)** properties to demonstrate the different effects that you can provide when entering a control and when dragging information from one control to another.

The sample uses two  **[TextBox](textbox-object-outlook-forms-script.md)** controls. You can set **DragBehavior** and **EnterFieldBehavior** for each control and see the effects of dragging from one control to another.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:


- A  **TextBox** named TextBox1.
    
- Two  **[ToggleButton](togglebutton-object-outlook-forms-script.md)** controls named ToggleButton1 and ToggleButton2. These controls are associated with TextBox1.
    
- A  **TextBox** named TextBox2.
    
- Two  **ToggleButton** controls named ToggleButton3 and ToggleButton4. These controls are associated with TextBox2.
    



```vb
Dim TextBox1, TextBox2 
Dim ToggleButton1, ToggleButton2, ToggleButton3, ToggleButton4 
 
Sub Item_Open() 
 set TextBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TextBox1") 
 set TextBox2 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TextBox2") 
 set ToggleButton1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ToggleButton1") 
 set ToggleButton2 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ToggleButton2") 
 set ToggleButton3 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ToggleButton3") 
 set ToggleButton4 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ToggleButton4") 
 
 TextBox1.Text = "Once upon a time in a land ...," 
 ToggleButton1.Value = True 
 ToggleButton1.Caption = "Drag Enabled" 
 ToggleButton1.WordWrap = True 
 TextBox1.DragBehavior = 1 'fmDragBehaviorEnabled 
 
 ToggleButton2.Value = True 
 ToggleButton2.Caption = "Recall Selection" 
 ToggleButton2.WordWrap = True 
 TextBox1.EnterFieldBehavior = 1 'fmEnterFieldBehaviorRecallSelection 
 
 TextBox2.Text = "XXX, YYYY" 
 ToggleButton3.Value = False 
 ToggleButton3.Caption = "Drag Disabled" 
 ToggleButton3.WordWrap = True 
 TextBox2.DragBehavior = 0 'fmDragBehaviorDisabled 
 
 ToggleButton4.Value = False 
 ToggleButton4.Caption = "Select All" 
 ToggleButton4.WordWrap = True 
 TextBox2.EnterFieldBehavior = 0 'fmEnterFieldBehaviorSelectAll 
End Sub 
 
Sub ToggleButton1_Click() 
 If ToggleButton1.Value = True Then 
 ToggleButton1.Caption = "Drag Enabled" 
 TextBox1.DragBehavior = 1 'fmDragBehaviorEnabled 
 Else 
 ToggleButton1.Caption = "Drag Disabled" 
 TextBox1.DragBehavior = 0 'fmDragBehaviorDisabled 
 End If 
End Sub 
 
Sub ToggleButton2_Click() 
 If ToggleButton2.Value = True Then 
 ToggleButton2.Caption = "Recall Selection" 
 TextBox1.EnterFieldBehavior = 1 'fmEnterFieldBehaviorRecallSelection 
 Else 
 ToggleButton2.Caption = "Select All" 
 TextBox1.EnterFieldBehavior = 0 'fmEnterFieldBehaviorSelectAll 
 End If 
End Sub 
 
Sub ToggleButton3_Click() 
 If ToggleButton3.Value = True Then 
 ToggleButton3.Caption = "Drag Enabled" 
 TextBox2.DragBehavior = 1 'fmDragBehaviorEnabled 
 Else 
 ToggleButton3.Caption = "Drag Disabled" 
 TextBox2.DragBehavior = 0 'fmDragBehaviorDisabled 
 End If 
End Sub 
 
Sub ToggleButton4_Click() 
 If ToggleButton4.Value = True Then 
 ToggleButton4.Caption = "Recall Selection" 
 TextBox2.EnterFieldBehavior = 1 'fmEnterFieldBehaviorRecallSelection 
 Else 
 ToggleButton4.Caption = "Select All" 
 TextBox2.EnterFieldBehavior = 0 'fmEnterFieldBehaviorSelectAll 
 End If 
End Sub
```


