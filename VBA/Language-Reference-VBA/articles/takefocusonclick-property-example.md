---
title: TakeFocusOnClick Property Example
keywords: fm20.chm5225119
f1_keywords:
- fm20.chm5225119
ms.prod: office
ms.assetid: fdc5a590-eee9-0ab2-aead-f3c02abf0eab
ms.date: 06/08/2017
---


# TakeFocusOnClick Property Example

The following example uses the  **TakeFocusOnClick** property to control whether a **CommandButton** receives the focus when the user clicks on it. The user clicks a control other than CommandButton1 and then clicks CommandButton1. If **TakeFocusOnClick** is **True**, CommandButton1 receives the focus after it is clicked. The user can change the value of **TakeFocusOnClick** by clicking the **ToggleButton**.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:




- A  **CommandButton** named CommandButton1.
    
- A  **ToggleButton** named ToggleButton1.
    
- One or two other controls, such as an  **OptionButton** or **ListBox**.
    




```vb
Private Sub CommandButton1_Click() 
 MsgBox "Watch CommandButton1 to see if it " _ 
 &; "takes the focus." 
End Sub 
 
Private Sub ToggleButton1_Click() 
 If ToggleButton1 = True Then 
 CommandButton1.TakeFocusOnClick = True 
 ToggleButton1.Caption = "TakeFocusOnClick On" 
 Else 
 CommandButton1.TakeFocusOnClick = False 
 ToggleButton1.Caption = "TakeFocusOnClick Off" 
 End If 
End Sub 
 
Private Sub UserForm_Initialize() 
 CommandButton1.Caption = "Show Message" 
 
 ToggleButton1.Caption = "TakeFocusOnClick On" 
 ToggleButton1.Value = True 
 ToggleButton1.Width = 90 
End Sub
```


