---
title: "How to: Place an Image and a Caption on a Command Button"
keywords: olfm10.chm3077228
f1_keywords:
- olfm10.chm3077228
ms.prod: outlook
ms.assetid: f17ced22-bbab-87e0-ef3a-5450c55fa950
ms.date: 06/08/2017
---


# How to: Place an Image and a Caption on a Command Button

The following example uses a  **[ComboBox](combobox-object-outlook-forms-script.md)** to show the picture placement options for a control. Each time the user clicks a list choice, the picture and caption are updated on the **[CommandButton](commandbutton-object-outlook-forms-script.md)**. This code sample also uses the  **[AddItem](combobox-additem-method-outlook-forms-script.md)** method to populate the **ComboBox** choices.

To use this example, copy this sample code to the Script Editor of a form. To run the code you need to open the form so the  **Open** event will activate. Make sure that the form contains:

- A  **[Label](label-object-outlook-forms-script.md)** named Label1.
    
- A  **CommandButton** named CommandButton1 with the **[Picture](commandbutton-picture-property-outlook-forms-script.md)** property set to use an image on your computer.
    
- A  **ComboBox** named ComboBox1.
    



```vb
Dim Label1 
Dim CommandButton1 
Dim ComboBox1 
 
Sub Item_Open() 
Set Label1 = Item.GetInspector.ModifiedFormPages("P.2").Label1 
Set CommandButton1 = Item.GetInspector.ModifiedFormPages("P.2").CommandButton1 
Set ComboBox1 = Item.GetInspector.ModifiedFormPages("P.2").ComboBox1 
 
 Label1.Left = 18 
 Label1.Top = 12 
 Label1.Height = 12 
 Label1.Width = 190 
 Label1.Caption = "Select picture placement relative to the caption." 
 
 'Add list entries to combo box. The value of each entry matches the 
 'corresponding ListIndex value in the combo box. 
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
 
 ComboBox1.Style = 2 'Use drop-down list 
 
 ComboBox1.BoundColumn = 0 'Combo box values are ListIndex values 
 ComboBox1.ListIndex = 0 'Set combo box to first entry 
 ComboBox1.Left = 18 
 ComboBox1.Top = 36 
 ComboBox1.Width = 90 
 ComboBox1.ListWidth = 90 
 
 'Initialize CommandButton1 
 CommandButton1.Left = 230 
 CommandButton1.Top = 36 
 CommandButton1.Height = 120 
 CommandButton1.Width = 120 
 
 'Note: Be sure to refer to have set the CommandButton1 to a bitmap file 
 'Note: that is present on your system 
 CommandButton1.PicturePosition = ComboBox1.Value 
End Sub 
 
Sub ComboBox1_Click() 
 Select Case ComboBox1.Value 
 Case 0 'Left Top 
 CommandButton1.Caption = "Left Top" 
 CommandButton1.PicturePosition = 0 
 
 Case 1 'Left Center 
 CommandButton1.Caption = "Left Center" 
 CommandButton1.PicturePosition = 1 
 
 Case 2 'Left Bottom 
 CommandButton1.Caption = "Left Bottom" 
 CommandButton1.PicturePosition = 2 
 
 Case 3 'Right Top 
 CommandButton1.Caption = "Right Top" 
 CommandButton1.PicturePosition = 3 
 
 Case 4 'Right Center 
 CommandButton1.Caption = "Right Center" 
 CommandButton1.PicturePosition = 4 
 
 Case 5 'Right Bottom 
 CommandButton1.Caption = "Right Bottom" 
 CommandButton1.PicturePosition = 5 
 
 Case 6 'Above Left 
 CommandButton1.Caption = "Above Left" 
 CommandButton1.PicturePosition = 6 
 
 Case 7 'Above Center 
 CommandButton1.Caption = "Above Center" 
 CommandButton1.PicturePosition = 7 
 
 Case 8 'Above Right 
 CommandButton1.Caption = "Above Right" 
 CommandButton1.PicturePosition = 8 
 
 Case 9 'Below Left 
 CommandButton1.Caption = "Below Left" 
 CommandButton1.PicturePosition = 9 
 
 Case 10 'Below Center 
 CommandButton1.Caption = "Below Center" 
 CommandButton1.PicturePosition = 10 
 
 Case 11 'Below Right 
 CommandButton1.Caption = "Below Right" 
 CommandButton1.PicturePosition = 11 
 
 Case 12 'Centered 
 CommandButton1.Caption = "Centered" 
 CommandButton1.PicturePosition = 12 
 
 End Select 
 
End Sub
```


