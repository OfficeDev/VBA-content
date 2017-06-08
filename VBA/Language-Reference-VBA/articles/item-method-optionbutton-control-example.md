---
title: Item Method, OptionButton Control Example
keywords: fm20.chm5225116
f1_keywords:
- fm20.chm5225116
ms.prod: office
ms.assetid: 1145cded-2cac-2631-9e7c-bed052283373
ms.date: 06/08/2017
---


# Item Method, OptionButton Control Example

The following example uses the  **Item** method to access individual members of the **Controls** and **Pages** collections. The user chooses an option button for either the **Controls** collection or the **MultiPage**, and then clicks the **CommandButton**. The name of the appropriate control is returned in the **Label**.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:




- A  **CommandButton** named CommandButton1.
    
- A  **Label** named Label1.
    
- Two  **OptionButton** controls named OptionButton1 and OptionButton2.
    
- A  **MultiPage** named MultiPage1.
    




```vb
Dim MyControl As Object 
Dim ControlsIndex As Integer 
 
Private Sub CommandButton1_Click() 
 If OptionButton1.Value = True Then 
 'Process Controls collection for UserForm 
 Set MyControl = Controls.Item(ControlsIndex) 
 Label1.Caption = MyControl.Name 
 
 'Prepare index for next control on Userform 
 ControlsIndex = ControlsIndex + 1 
 If ControlsIndex >= Controls.Count Then 
 ControlsIndex = 0 
 End If 
 
 ElseIf OptionButton2.Value = True Then 
 'Process Current Page of Pages collection 
 Set MyControl = MultiPage1.Pages _ 
 .Item(MultiPage1.Value) 
 Label1.Caption = MyControl.Name 
 End If 
End Sub 
 
Private Sub UserForm_Initialize() 
 ControlsIndex = 0 
 
 OptionButton1.Caption = "Controls Collection" 
 OptionButton2.Caption = "Pages Collection" 
 OptionButton1.Value = True 
 
 CommandButton1.Caption = "Get Member Name" 
End Sub
```


