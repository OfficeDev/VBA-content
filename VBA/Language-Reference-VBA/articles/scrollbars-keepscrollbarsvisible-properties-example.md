---
title: ScrollBars, KeepScrollBarsVisible Properties Example
keywords: fm20.chm5225137
f1_keywords:
- fm20.chm5225137
ms.prod: office
ms.assetid: a935d8ab-2060-2794-69a8-ba7c8ceed3d1
ms.date: 06/08/2017
---


# ScrollBars, KeepScrollBarsVisible Properties Example

The following example uses the  **ScrollBars** and the **KeepScrollBarsVisible** properties to add scroll bars to a page of a **MultiPage** and to a **Frame**. The user chooses an option button that, in turn, specifies a value for **KeepScrollBarsVisible**.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:




- A  **MultiPage** named MultiPage1.
    
- A  **Frame** named Frame1.
    
- Four  **OptionButton** controls named OptionButton1 through OptionButton4.
    




```vb
Private Sub UserForm_Initialize() 
 MultiPage1.Pages(0).ScrollBars = fmScrollBarsBoth 
 MultiPage1.Pages(0).KeepScrollBarsVisible = fmScrollBarsNone 
 
 Frame1.ScrollBars = fmScrollBarsBoth 
 Frame1.KeepScrollBarsVisible = fmScrollBarsNone 
 
 OptionButton1.Caption = "No scroll bars" 
 OptionButton1.Value = True 
 OptionButton2.Caption = "Horizontal scroll bars" 
 OptionButton3.Caption = "Vertical scroll bars" 
 OptionButton4.Caption = "Both scroll bars" 
End Sub 
 
Private Sub OptionButton1_Click() 
 MultiPage1.Pages(0).KeepScrollBarsVisible = _ 
 fmScrollBarsNone 
 Frame1.KeepScrollBarsVisible = fmScrollBarsNone 
End Sub 
 
Private Sub OptionButton2_Click() 
 MultiPage1.Pages(0).KeepScrollBarsVisible = _ 
 fmScrollBarsHorizontal 
 Frame1.KeepScrollBarsVisible = _ 
 fmScrollBarsHorizontal 
End Sub 
 
Private Sub OptionButton3_Click() 
 MultiPage1.Pages(0).KeepScrollBarsVisible = _ 
 fmScrollBarsVertical 
 Frame1.KeepScrollBarsVisible = _ 
 fmScrollBarsVertical 
End Sub 
 
Private Sub OptionButton4_Click() 
 MultiPage1.Pages(0).KeepScrollBarsVisible = _ 
 fmScrollBarsBoth 
 Frame1.KeepScrollBarsVisible = fmScrollBarsBoth 
End Sub
```


