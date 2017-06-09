---
title: "How to: Use the Controls Collection to Move Controls on a Page"
keywords: olfm10.chm3077149
f1_keywords:
- olfm10.chm3077149
ms.prod: outlook
ms.assetid: 19170632-76c6-3ca9-d7ea-f68323d878a6
ms.date: 06/08/2017
---


# How to: Use the Controls Collection to Move Controls on a Page

The following example accesses individual controls from the Microsoft Forms 2.0  **Controls** collection using a `For Each...Next` loop. When the user presses CommandButton1, the other controls are placed in a column along the left edge of the form using the **Move** method of the control.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event of the item will activate. Make sure that the form contains a **[CommandButton](commandbutton-object-outlook-forms-script.md)** named CommandButton1 and several other controls.



```vb
Dim CtrlHeight 
Dim CtrlTop 
Dim CtrlGap 
Dim CommandButton1 
 
Sub Item_Open() 
 Set CommandButton1 = Item.GetInspector.ModifiedFormPages("P.2").CommandButton1 
 
 CtrlHeight = 20 
 CtrlGap = 5 
 
 CommandButton1.Caption = "Click to move controls" 
 CommandButton1.AutoSize = True 
 CommandButton1.Left = 120 
 CommandButton1.Top = CtrlTop 
End Sub 
 
Sub CommandButton1_Click() 
 Dim MyControl 
 
 Set AllControls = Item.GetInspector.ModifiedFormPages("P.2").Controls 
 
 CtrlTop = 5 
 
 For i = 0 to AllControls.Count - 1 
 Set MyControl = AllControls(i) 
 If MyControl.Name = "CommandButton1" Then 
 'Don't move or resize this control. 
 Else 
 'Move method using unnamed arguments (left, top, width, height) 
 MyControl.Move 5, CtrlTop, ,CtrlHeight 
 
 'Calculate top coordinate for next control 
 CtrlTop = CtrlTop + CtrlHeight + CtrlGap 
 End If 
 Next 
 
End Sub
```


