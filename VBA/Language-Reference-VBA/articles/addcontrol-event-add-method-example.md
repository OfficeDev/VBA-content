---
title: AddControl Event, Add Method Example
keywords: fm20.chm5225176
f1_keywords:
- fm20.chm5225176
ms.prod: office
ms.assetid: 6a57bc57-7971-c6b1-72a1-78d5c835b380
ms.date: 06/08/2017
---


# AddControl Event, Add Method Example

The following example uses the  **Add** method to add a control to a form at run time and uses the **AddControl** event as confirmation that the control was added.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:




- A  **CommandButton** named CommandButton1.
    
- A  **Label** named Label1.
    




```vb
Dim Mycmd as Control 
Private Sub CommandButton1_Click() 
 
 Set Mycmd = Controls.Add("MSForms.CommandButton.1") ', CommandButton2, Visible) 
 Mycmd.Left = 18 
 Mycmd.Top = 150 
 Mycmd.Width = 175 
 Mycmd.Height = 20 
 Mycmd.Caption = "This is fun." &; Mycmd.Name 
 
End Sub 
 
Private Sub UserForm_AddControl(ByVal Control As _ 
 MSForms.Control) 
 Label1.Caption = "Control was Added." 
End Sub
```


