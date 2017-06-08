---
title: Page Object, MultiPage Control, Add, Clear, Remove Methods Example
keywords: fm20.chm5225177
f1_keywords:
- fm20.chm5225177
ms.prod: office
ms.assetid: ba40e297-6f1f-b012-34a2-d8e6c6b0e462
ms.date: 06/08/2017
---


# Page Object, MultiPage Control, Add, Clear, Remove Methods Example

The following example uses the  **Add**, **Clear**, and **Remove** methods to add and remove a control to a **Page** of a **MultiPage** at run time.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:




- A  **MultiPage** named MultiPage1.
    
- Three  **CommandButton** controls named CommandButton1 through CommandButton3.
    




```vb
Dim MyTextBox As Control 
 
Private Sub CommandButton1_Click() 
Set MyTextBox = MultiPage1.Pages(0).Controls.Add("MSForms" _ 
 &; ".TextBox.1", "MyTextBox", Visible) 
End Sub 
 
Private Sub CommandButton2_Click() 
 MultiPage1.Pages(0).Controls.Clear 
End Sub 
 
Private Sub CommandButton3_Click() 
 If MultiPage1.Pages(0).Controls.Count > 0 Then 
 MultiPage1.Pages(0).Controls.Remove "MyTextBox" 
 End If 
End Sub 
 
Private Sub UserForm_Initialize() 
 CommandButton1.Caption = "Add control" 
 CommandButton2.Caption = "Clear controls" 
 CommandButton3.Caption = "Remove control" 
End Sub
```


