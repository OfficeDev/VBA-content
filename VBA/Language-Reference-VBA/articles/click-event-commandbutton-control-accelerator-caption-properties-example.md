---
title: Click Event, CommandButton Control, Accelerator, Caption Properties Example
keywords: fm20.chm5225180
f1_keywords:
- fm20.chm5225180
ms.prod: office
ms.assetid: f2d2210a-e69e-6dbb-6b3d-95ceb377bc84
ms.date: 06/08/2017
---


# Click Event, CommandButton Control, Accelerator, Caption Properties Example

This example changes the  **Accelerator** and **Caption** properties of a **CommandButton** each time the user clicks the button by using the mouse or the accelerator key. The Click event contains the code to change the **Accelerator** and **Caption** properties.

To try this example, paste the code into the Declarations section of a form containing a  **CommandButton** named CommandButton1.



```vb
Private Sub UserForm_Initialize() 
 CommandButton1.Accelerator = "C" 
 'Set Accelerator key to COMMAND + C 
End Sub 
 
Private Sub CommandButton1_Click () 
 If CommandButton1.Caption = "OK" Then 
 'Check caption, then change it. 
 CommandButton1.Caption = "Clicked" 
 CommandButton1.Accelerator = "C" 
 'Set Accelerator key to COMMAND + C 
 Else 
 CommandButton1.Caption = "OK" 
 CommandButton1.Accelerator = "O" 
 'Set Accelerator key to COMMAND + O 
 End If 
End Sub
```


