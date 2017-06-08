---
title: Zoom Event, Zoom Property, Label Control Example
keywords: fm20.chm5225134
f1_keywords:
- fm20.chm5225134
ms.prod: office
ms.assetid: 1ded265c-6682-221f-e3c3-1ebf08a550c0
ms.date: 06/08/2017
---


# Zoom Event, Zoom Property, Label Control Example

The following example uses the  **Zoom** event to evaluate the new value of the **Zoom** property and adds scroll bars to the form when appropriate. The example uses a **Label** to display the current value. The user specifies the size for the form by using the **SpinButton** and then clicks the **CommandButton** to set the value in the **Zoom** property.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:




- A  **Label** named Label1.
    
- A  **SpinButton** named SpinButton1.
    
- A  **CommandButton** named CommandButton1.
    
- Other controls placed near the edges of the form.
    




```vb
Private Sub CommandButton1_Click() 
 Zoom = SpinButton1.Value 
End Sub 
 
Private Sub SpinButton1_SpinDown() 
 Label1.Caption = SpinButton1.Value 
End Sub 
 
Private Sub SpinButton1_SpinUp() 
 Label1.Caption = SpinButton1.Value 
End Sub 
 
Private Sub UserForm_Initialize() 
 SpinButton1.Min = 10 
 SpinButton1.Max = 400 
 SpinButton1.Value = 100 
 Label1.Caption = SpinButton1.Value 
 
 CommandButton1.Caption = "Zoom it!" 
End Sub 
 
Private Sub UserForm_Zoom(Percent As Integer) 
 Dim MyResult As Double 
 
 If Percent > 99 Then 
 ScrollBars = fmScrollBarsBoth 
 ScrollLeft = 0 
 ScrollTop = 0 
 
 MyResult = Width * Percent / 100 
 ScrollWidth = MyResult 
 
 MyResult = Height * Percent / 100 
 ScrollHeight = MyResult 
 Else 
 ScrollBars = fmScrollBarsNone 
 ScrollLeft = 0 
 ScrollTop = 0 
 End If 
End Sub
```


