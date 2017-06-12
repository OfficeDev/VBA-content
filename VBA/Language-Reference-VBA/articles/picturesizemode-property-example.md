---
title: PictureSizeMode Property Example
keywords: fm20.chm5225160
f1_keywords:
- fm20.chm5225160
ms.prod: office
ms.assetid: 8a535e7b-e32c-9bd4-1ee8-9560fb01c4d9
ms.date: 06/08/2017
---


# PictureSizeMode Property Example

The following example uses the  **PictureSizeMode** property to demonstrate three display options for a picture: showing the picture as is, changing the size of the picture while maintaining its original proportions, and stretching the picture to fill a space.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:




- A  **Frame** named Frame1.
    
- A  **SpinButton** named SpinButton1.
    
- A  **TextBox** named TextBox1.
    
- Three  **OptionButton** controls named OptionButton1 through OptionButton3.
    


 **Note**  This example is an enhanced version of the  **PictureAlignment** property example, as the two properties complement each other. The enhancements are three **OptionButton** event subroutines that control whether the image is cropped, zoomed, or stretched.




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
 
 OptionButton1.Caption = "Crop" 
 OptionButton1.Value = True 
 OptionButton2.Caption = "Stretch" 
 OptionButton3.Caption = "Zoom" 
End Sub 
 
Private Sub OptionButton1_Click() 
 If OptionButton1.Value = True Then 
 Frame1.PictureSizeMode = fmPictureSizeModeClip 
 End If 
End Sub 
 
Private Sub OptionButton2_Click() 
 If OptionButton2.Value = True Then 
 Frame1.PictureSizeMode = fmPictureSizeModeStretch 
 End If 
End Sub 
 
Private Sub OptionButton3_Click() 
 If OptionButton3.Value = True Then 
 Frame1.PictureSizeMode = fmPictureSizeModeZoom 
 End If 
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


