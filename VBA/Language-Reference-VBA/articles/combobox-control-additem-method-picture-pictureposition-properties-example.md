---
title: ComboBox Control, AddItem Method, Picture, PicturePosition Properties Example
keywords: fm20.chm5225189
f1_keywords:
- fm20.chm5225189
ms.prod: office
ms.assetid: 6af29fd2-49f5-980a-b72e-9776a82e3170
ms.date: 06/08/2017
---


# ComboBox Control, AddItem Method, Picture, PicturePosition Properties Example

The following example uses a  **ComboBox** to show the picture placement options for a control. Each time the user clicks a list choice, the picture and caption are updated on the **CommandButton**. This code sample also uses the **AddItem** method to populate the **ComboBox** choices.

To use this example, copy this sample code to the Declarations portion of a form. Make sure that the form contains:




- A  **Label** named Label1.
    
- A  **CommandButton** named CommandButton1.
    
- A  **ComboBox** named ComboBox1.
    




```vb
Private Sub UserForm_Initialize() 
 Label1.Left = 18 
 Label1.Top = 12 
 Label1.Height = 12 
 Label1.Width = 190 
 Label1.Caption = "Select picture placement " _ 
 &; "relative to the caption." 
 
 'Add list entries to combo box. The value of each 
 'entry matches the corresponding ListIndex value 
 'in the combo box. 
 ComboBox1.AddItem "Left Top" 'ListIndex = 0 
 ComboBox1.AddItem "Left Center" 'ListIndex = 1 
 ComboBox1.AddItem "Left Bottom" 'ListIndex = 2 
 ComboBox1.AddItem "Right Top" 'ListIndex = 3 
 ComboBox1.AddItem "Right Center" 'ListIndex = 4 
 ComboBox1.AddItem "Right Bottom" 'ListIndex = 5 
 ComboBox1.AddItem "Above Left" 'ListIndex = 6 
 ComboBox1.AddItem "Above Center" 'ListIndex = 7 
 ComboBox1.AddItem "Above Right" 'ListIndex = 8 
 ComboBox1.AddItem "Below Left" 'ListIndex = 9 
 ComboBox1.AddItem "Below Center" 'ListIndex = 10 
 ComboBox1.AddItem "Below Right" 'ListIndex = 11 
 ComboBox1.AddItem "Centered" 'ListIndex = 12 
 'Use drop-down list 
 ComboBox1.Style = fmStyleDropDownList 
 'Combo box values are ListIndex values 
 ComboBox1.BoundColumn = 0 
 'Set combo box to first entry 
 ComboBox1.ListIndex = 0 
 
 
 ComboBox1.Left = 18 
 ComboBox1.Top = 36 
 ComboBox1.Width = 90 
 ComboBox1.ListWidth = 90 
 
 'Initialize CommandButton1 
 CommandButton1.Left = 230 
 CommandButton1.Top = 36 
 CommandButton1.Height = 120 
 CommandButton1.Width = 120 
 
 'Note: Be sure to refer to a bitmap file that is 
 'present on your system, and to include the path 
 'in the filename. 
 CommandButton1.Picture = _ 
 LoadPicture("c:\windows\argyle.bmp") 
 CommandButton1.PicturePosition = ComboBox1.Value 
End Sub 
 
Private Sub ComboBox1_Click() 
 Select Case ComboBox1.Value 
 Case 0 'Left Top 
 CommandButton1.Caption = "Left Top" 
 CommandButton1.PicturePosition = _ 
 fmPicturePositionLeftTop 
 
 Case 1 'Left Center 
 CommandButton1.Caption = "Left Center" 
 CommandButton1.PicturePosition = _ 
 fmPicturePositionLeftCenter 
 
 Case 2 'Left Bottom 
 CommandButton1.Caption = "Left Bottom" 
 CommandButton1.PicturePosition = _ 
 fmPicturePositionLeftBottom 
 
 Case 3 'Right Top 
 CommandButton1.Caption = "Right Top" 
 CommandButton1.PicturePosition = _ 
 fmPicturePositionRightTop 
 
 Case 4 'Right Center 
 CommandButton1.Caption = "Right Center" 
 CommandButton1.PicturePosition = _ 
 fmPicturePositionRightCenter 
 
 Case 5 'Right Bottom 
 CommandButton1.Caption = "Right Bottom" 
 CommandButton1.PicturePosition = _ 
 fmPicturePositionRightBottom 
 
 Case 6 'Above Left 
 CommandButton1.Caption = "Above Left" 
 CommandButton1.PicturePosition = _ 
 fmPicturePositionAboveLeft 
 
 Case 7 'Above Center 
 CommandButton1.Caption = "Above Center" 
 CommandButton1.PicturePosition = _ 
 fmPicturePositionAboveCenter 
 
 Case 8 'Above Right 
 CommandButton1.Caption = "Above Right" 
 CommandButton1.PicturePosition = _ 
 fmPicturePositionAboveRight 
 
 Case 9 'Below Left 
 CommandButton1.Caption = "Below Left" 
 CommandButton1.PicturePosition = _ 
 fmPicturePositionBelowLeft 
 
 Case 10 'Below Center 
 CommandButton1.Caption = "Below Center" 
 CommandButton1.PicturePosition = _ 
 fmPicturePositionBelowCenter 
 
 Case 11 'Below Right 
 CommandButton1.Caption = "Below Right" 
 CommandButton1.PicturePosition = _ 
 fmPicturePositionBelowRight 
 
 Case 12 'Centered 
 CommandButton1.Caption = "Centered" 
 CommandButton1.PicturePosition = _ 
 fmPicturePositionCenter 
 
 End Select 
 
End Sub
```


