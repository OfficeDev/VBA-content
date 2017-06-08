---
title: PictureAlignment Property Example
keywords: fm20.chm5225161
f1_keywords:
- fm20.chm5225161
ms.prod: office
ms.assetid: 92ddb3be-c005-7fb3-dbfd-2cbd74ca021c
ms.date: 06/08/2017
---


# PictureAlignment Property Example

The following example uses the  **PictureAlignment** property to set up a background picture. The example also identifies the alignment options provided by **PictureAlignment**.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:




- A  **Frame** named Frame1.
    
- A  **SpinButton** named SpinButton1.
    
- A  **TextBox** named TextBox1.
    




```vb
Dim Alignments(5) As String 
 
Private Sub UserForm_Initialize() 
 Alignments(0) = "0 - Top Left" 
 Alignments(1) = "1 - Top Right" 
 Alignments(2) = "2 - Center" 
 Alignments(3) = "3 - Bottom Left" 
 Alignments(4) = "4 - Bottom Right" 
 
 'Specify a bitmap that exists on your system 
 Frame1.Picture = LoadPicture("c:\winnt2\ball.bmp") 
 
 SpinButton1.Min = 0 
 SpinButton1.Max = 4 
 SpinButton1.Value = 0 
 
 TextBox1.Text = Alignments(0) 
 Frame1.PictureAlignment = SpinButton1.Value 
End Sub 
 
Private Sub SpinButton1_Change() 
 TextBox1.Text = Alignments(SpinButton1.Value) 
 Frame1.PictureAlignment = SpinButton1.Value 
End Sub 
 
Private Sub TextBox1_Change() 
 Select Case TextBox1.Text 
 Case "0" 
 TextBox1.Text = Alignments(0) 
 Frame1.PictureAlignment = 0 
 Case "1" 
 TextBox1.Text = Alignments(1) 
 Frame1.PictureAlignment = 1 
 Case "2" 
 TextBox1.Text = Alignments(2) 
 Frame1.PictureAlignment = 2 
 Case "3" 
 TextBox1.Text = Alignments(3) 
 Frame1.PictureAlignment = 3 
 Case "4" 
 TextBox1.Text = Alignments(4) 
 Frame1.PictureAlignment = 4 
 Case Else 
 TextBox1.Text = Alignments(SpinButton1.Value) 
 Frame1.PictureAlignment = SpinButton1.Value 
 End Select 
End Sub
```


