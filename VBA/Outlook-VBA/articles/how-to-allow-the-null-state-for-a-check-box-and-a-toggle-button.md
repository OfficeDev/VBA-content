---
title: "How to: Allow the Null State for a Check Box and a Toggle Button"
keywords: olfm10.chm3077261
f1_keywords:
- olfm10.chm3077261
ms.prod: outlook
ms.assetid: 75b3374d-6d96-3bcc-3e97-f0089f3fdd99
ms.date: 06/08/2017
---


# How to: Allow the Null State for a Check Box and a Toggle Button

The following example uses the  **TripleState** property to allow Null as a legal value of a **[CheckBox](checkbox-object-outlook-forms-script.md)** and a **[ToggleButton](togglebutton-object-outlook-forms-script.md)**. The user controls the value of  **TripleState** through ToggleButton2. The user can set the value of a **CheckBox** or **ToggleButton** based on the value of **TripleState**. However, when a control is set to  **Null**, no event is fired.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:

- A  **CheckBox** named CheckBox1.
    
- A  **ToggleButton** named ToggleButton1.
    
- A  **ToggleButton** named ToggleButton2.
    



```vb
Sub Item_Open() 
 Set CheckBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("CheckBox1") 
 Set ToggleButton1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ToggleButton1") 
 Set ToggleButton2 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ToggleButton2") 
 
 CheckBox1.Caption = "Value is True" 
 CheckBox1.Value = True 
 CheckBox1.TripleState = False 
 
 ToggleButton1.Caption = "Value is True" 
 ToggleButton1.Value = True 
 ToggleButton1.TripleState = False 
 
 ToggleButton2.Value = False 
 ToggleButton2.Caption = "Triple State Off" 
End Sub 
 
Sub ToggleButton2_Click() 
 Set CheckBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("CheckBox1") 
 Set ToggleButton1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ToggleButton1") 
 Set ToggleButton2 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ToggleButton2") 
 
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
 
Sub CheckBox1_Click() 
 Set CheckBox1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("CheckBox1") 
 If IsNull(CheckBox1.Value) Then 
 CheckBox1.Caption = "Value is Null" 
 ElseIf CheckBox1.Value = False Then 
 CheckBox1.Caption = "Value is False" 
 ElseIf CheckBox1.Value = True Then 
 CheckBox1.Caption = "Value is True" 
 End If 
End Sub 
 
Sub ToggleButton1_Click() 
 Set ToggleButton1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ToggleButton1") 
 If IsNull(ToggleButton1.Value) Then 
 ToggleButton1.Caption = "Value is Null" 
 ElseIf ToggleButton1.Value = False Then 
 ToggleButton1.Caption = "Value is False" 
 ElseIf ToggleButton1.Value = True Then 
 ToggleButton1.Caption = "Value is True" 
 End If 
End Sub
```


