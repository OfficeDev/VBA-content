---
title: ScrollBar Control, Scroll Event Example
keywords: fm20.chm5225168
f1_keywords:
- fm20.chm5225168
ms.prod: office
ms.assetid: e4180c7f-d14e-b76e-fd7a-b1cf354b0fd0
ms.date: 06/08/2017
---


# ScrollBar Control, Scroll Event Example

The following example demonstrates the stand-alone  **ScrollBar** and reports the change in its value as the user moves the scroll box. The user can move the scroll box by clicking on either arrow at the ends of the control, by clicking in the region between scroll box and either arrow, or by dragging the scroll box. When the user drags the scroll box, the Scroll event displays a message indicating that the user scrolled to obtain the new value.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:




- A  **ScrollBar** named ScrollBar1.
    
- Two  **Label** controls named Label1 and Label2. Label1 contains scaling information for the user. Label2 reports the delta value.
    




```vb
Dim ScrollSaved As Integer 
'Previous ScrollBar setting 
 
Private Sub UserForm_Initialize() 
 ScrollBar1.Min = -225 
 ScrollBar1.Max = 289 
 ScrollBar1.Value = 0 
 
 Label1.Caption = "-225 -----Widgets----- 289" 
 Label1.AutoSize = True 
 
 Label2.Caption = "" 
End Sub 
 
Private Sub ScrollBar1_Change() 
 Label2.Caption = " Widget Changes " _ 
 &; (ScrollSaved - ScrollBar1.Value) 
End Sub 
 
Private Sub ScrollBar1_Exit(ByVal Cancel as MSForms.ReturnBoolean) 
 Label2.Caption = " Widget Changes " _ 
 &; (ScrollSaved - ScrollBar1.Value) 
 ScrollSaved = ScrollBar1.Value 
End Sub 
 
Private Sub ScrollBar1_Scroll() 
 Label2.Caption = (ScrollSaved - ScrollBar1 _ 
 .Value) &; " Widget Changes by Scrolling" 
End Sub
```


