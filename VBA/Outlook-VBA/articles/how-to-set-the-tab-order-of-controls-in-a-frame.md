---
title: "How to: Set the Tab Order of Controls in a Frame"
keywords: olfm10.chm3077250
f1_keywords:
- olfm10.chm3077250
ms.prod: outlook
ms.assetid: 6525530b-e9a3-4285-30c5-0b9dd0e289d8
ms.date: 06/08/2017
---


# How to: Set the Tab Order of Controls in a Frame

The following example uses the  **TabIndex** property to display and set the tab order for individual controls. The **TabIndex** property is a Microsoft Forms 2.0 property that applies to every control that can exist in a **[Frame](frame-object-outlook-forms-script.md)**. The user can press TAB to reach the next control in the tab order and to display the  **TabIndex** of that control. The user can also click on any control, except a **[TextBox](textbox-object-outlook-forms-script.md)** or **[ScrollBar](scrollbar-object-outlook-forms-script.md)**, to display its  **TabIndex**. The user can change the  **TabIndex** of a control by specifying a new index value in the **TextBox** and clicking CommandButton3. Changing the **TabIndex** for one control also updates the **TabIndex** for other controls in the **Frame**.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:

- A  **[Label](label-object-outlook-forms-script.md)** named Label1.
    
- A  **TextBox** named TextBox1.
    
- A  **Frame** named Frame1.
    
- A  **TextBox** in the **Frame** named TextBox2.
    
- Two  **[CommandButton](commandbutton-object-outlook-forms-script.md)** controls in the **Frame** named CommandButton1 and CommandButton2.
    
- A  **ScrollBar** in the **Frame** named ScrollBar1.
    
- A  **CommandButton** (not in the **Frame**) named CommandButton3.
    



```vb
Sub MoveToFront() 
 Set Frame1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("Frame1") 
 Set TextBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TextBox1") 
 
 Temp = Frame1.ActiveControl.TabIndex 
 For i = 0 To Temp - 1 
 Frame1.Controls.Item(i).TabIndex = i + 1 
 Next 
 
 Frame1.ActiveControl.TabIndex = 0 
 TextBox1.Text = Frame1.ActiveControl.TabIndex 
End Sub 
 
Sub CommandButton3_Click() 
 Set Frame1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("Frame1") 
 Set TextBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TextBox1") 
 
 If IsNumeric(TextBox1.Text) Then 
 Temp = CInt(TextBox1.Text) 
 
 If Temp >= Frame1.Controls.Count Or Temp < 0 Then 
 'Entry out of range; move control to front of tab order 
 MoveToFront 
 ElseIf Temp > Frame1.ActiveControl.TabIndex Then 
 'Move entry down the list 
 For i = Frame1.ActiveControl.TabIndex + 1 To Temp 
 Frame1.Controls.Item(i).TabIndex = i - 1 
 Next 
 Frame1.ActiveControl.TabIndex = Temp 
 TextBox1.Text = Frame1.ActiveControl.TabIndex 
 Else 
 'Move Entry up the list 
 For i = Frame1.ActiveControl.TabIndex - 1 To Temp 
 Frame1.Controls.Item(i).TabIndex = i + 1 
 Next 
 Frame1.ActiveControl.TabIndex = Temp 
 TextBox1.Text = Frame1.ActiveControl.TabIndex 
 End If 
 Else 
 'Text entry; move control to front of tab order 
 MoveToFront 
 End If 
End Sub 
 
Sub Item_Open() 
 Set Frame1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("Frame1") 
 Set TextBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TextBox1") 
 Set Label1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("Label1") 
 Set CommandButton3 = Item.GetInspector.ModifiedFormPages("P.2").Controls("CommandButton3") 
 
 Label1.Caption = "TabIndex" 
 
 Frame1.Controls(0).SetFocus 
 TextBox1.Text = Frame1.ActiveControl.TabIndex 
 
 Frame1.Cycle = 2 '2=fmCycleCurrentForm 
 
 CommandButton3.Caption = "Set TabIndex" 
 CommandButton3.TakeFocusOnClick = False 
End Sub 
 
Sub CommandButton1_Click() 
 Set Frame1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("Frame1") 
 Set TextBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TextBox1") 
 
 TextBox1.Text = Frame1.ActiveControl.TabIndex 
End Sub 
 
Sub CommandButton2_Click() 
 Set Frame1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("Frame1") 
 Set TextBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("TextBox1") 
 
 TextBox1.Text = Frame1.ActiveControl.TabIndex 
End Sub
```


