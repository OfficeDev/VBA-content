---
title: "How to: Specify Tab Support for a Control"
keywords: olfm10.chm3077251
f1_keywords:
- olfm10.chm3077251
ms.prod: outlook
ms.assetid: e4b453a2-a0e2-63f0-3d93-a46e842fbbd6
ms.date: 06/08/2017
---


# How to: Specify Tab Support for a Control

The following example uses the  **TabStop** property to control whether a user can press TAB to move the focus to a particular control. The **TabIndex** property is a Microsoft Forms 2.0 property that applies to every control that supports tabbing. The user presses TAB to move the focus among the controls on the form, and then clicks the **[ToggleButton](togglebutton-object-outlook-forms-script.md)** to change **TabStop** for CommandButton1. When **TabStop** is **False**, CommandButton1 will not receive the focus by using TAB.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:

- A  **[CommandButton](commandbutton-object-outlook-forms-script.md)** named CommandButton1.
    
- A  **ToggleButton** named ToggleButton1.
    
- One or two other controls, such as an  **[OptionButton](optionbutton-object-outlook-forms-script.md)** or **[ListBox](listbox-object-outlook-forms-script.md)**.
    



```vb
Sub CommandButton1_Click() 
 MsgBox "Clicked CommandButton1." 
End Sub 
 
Sub ToggleButton1_Click() 
 Dim CommandButton1 
 Set ToggleButton1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ToggleButton1") 
 Set CommandButton1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("CommandButton1") 
 
 If ToggleButton1 = True Then 
 CommandButton1.TabStop = True 
 ToggleButton1.Caption = "TabStop On" 
 Else 
 CommandButton1.TabStop = False 
 ToggleButton1.Caption = "TabStop Off" 
 End If 
End Sub 
 
Sub Item_Open() 
 Set ToggleButton1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("ToggleButton1") 
 Set CommandButton1 = Item.GetInspector.ModifiedFormPages("P.2").Controls("CommandButton1") 
 
 CommandButton1.Caption = "Show Message" 
 
 ToggleButton1.Caption = "TabStop On" 
 ToggleButton1.Value = True 
 ToggleButton1.Width = 90 
End Sub
```


