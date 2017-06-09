---
title: Controls Collection, Move Method Example
keywords: fm20.chm5225192
f1_keywords:
- fm20.chm5225192
ms.prod: office
ms.assetid: 14694f03-8d28-9808-b413-96555f0fbc4b
ms.date: 06/08/2017
---


# Controls Collection, Move Method Example

The following example accesses individual controls from the  **Controls** collection using a **For Each...Next** loop. When the user presses CommandButton1, the other controls are placed in a column along the left edge of the form using the **Move** method.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains a  **CommandButton** named CommandButton1 and several other controls.



```vb
Dim CtrlHeight As Single 
Dim CtrlTop As Single 
Dim CtrlGap As Single 
 
Private Sub CommandButton1_Click() 
 Dim MyControl As Control 
 CtrlTop = 5 
 
 For Each MyControl In Controls 
 If MyControl.Name = "CommandButton1" Then 
 'Don't move or resize this control. 
 Else 
 'Move method using named arguments 
 MyControl.Move Top:=CtrlTop, _ 
 Height:=CtrlHeight, Left:=5 
 
 'Move method using unnamed arguments (left, 
 'top, width, height) 
 'MyControl.Move 5, CtrlTop, ,CtrlHeight 
 
 'Calculate top coordinate for next control 
 CtrlTop = CtrlTop + CtrlHeight + CtrlGap 
 End If 
 Next 
 
End Sub
```




```vb
Private Sub UserForm_Initialize() 
 CtrlHeight = 20 
 CtrlGap = 5 
 
 CommandButton1.Caption = "Click to move controls" 
 CommandButton1.AutoSize = True 
 CommandButton1.Left = 120 
 CommandButton1.Top = CtrlTop 
End Sub
```


