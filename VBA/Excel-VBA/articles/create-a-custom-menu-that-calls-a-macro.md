---
title: Create a Custom Menu That Calls a Macro
ms.prod: excel
ms.assetid: 925976ab-e2ef-4b71-aa06-62fe6ac8a4c3
ms.date: 06/08/2017
---


# Create a Custom Menu That Calls a Macro

The following code example shows how to create a custom menu with four menu options, each of which calls a macro.

 **Sample code provided by:** Holy Macro! Books, [Holy Macro! It's 2,500 Excel VBA Examples](http://www.mrexcel.com/store/index.php?l=product_detail&;p=1)

The following code example sets up the custom menu when the workbook is opened, and deletes it when the workbook is closed.




```vb
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
   With Application.CommandBars("Worksheet Menu Bar")
      On Error Resume Next
      .Controls("&;MyFunction").Delete
      On Error GoTo 0
   End With
End Sub

Private Sub Workbook_Open()
   Dim objPopUp As CommandBarPopup
   Dim objBtn As CommandBarButton
   With Application.CommandBars("Worksheet Menu Bar")
      On Error Resume Next
      .Controls("MyFunction").Delete
      On Error GoTo 0
      Set objPopUp = .Controls.Add( _
         Type:=msoControlPopup, _
         before:=.Controls.Count, _
         temporary:=True)
   End With
   objPopUp.Caption = "&;MyFunction"
   Set objBtn = objPopUp.Controls.Add
   With objBtn
      .Caption = "Formula Entry"
      .OnAction = "Cbm_Active_Formula"
      .Style = msoButtonCaption
   End With
   Set objBtn = objPopUp.Controls.Add
   With objBtn
      .Caption = "Value Entry"
      .OnAction = "Cbm_Active_Value"
      .Style = msoButtonCaption
   End With
   Set objBtn = objPopUp.Controls.Add
   With objBtn
      .Caption = "Formula Selection"
      .OnAction = "Cbm_Formula_Select"
      .Style = msoButtonCaption
   End With
   Set objBtn = objPopUp.Controls.Add
   With objBtn
      .Caption = "Value Selection"
      .OnAction = "Cbm_Value_Select"
      .Style = msoButtonCaption
   End With
End Sub
```

The menu "MyFunction" is added when the workbook opens, and is deleted when the workbook closes. It provides four menu options, with a macro assigned to each option. The user-defined function (UDF) "MyFunction" multiplies three values in a range together and returns the result.



```vb
Function MyFunction(rng As Range) As Double
   MyFunction = rng(1) * rng(2) * rng(3)
End Function
```

 **Formula Entry:** This menu option is assigned the macro "Cbm_Active_Formula", which calls the UDF named "MyFunction" that multiplies the numbers in the preceding 3 cells, and stores value of the UDF in the active cell. You must have values in the range B6:D6 and select cell E6 before clicking this menu option.



```vb
Sub Cbm_Active_Formula()
   'setting up the variables.
   Dim intLen As Integer, strRng As String
   
   'Using the active cell, E6.
   With ActiveCell
      'Check to see if the preceding offset has valid data, and if there are three values
      If IsEmpty(.Offset(0, -1)) Or .Column < 4 Then
         
          'If the data is not valid, call MyFunction directly as a formula, but with no parameters.
         .Formula = "=MyFunction()"
          Application.SendKeys "{ENTER}"
      Else
      
         'If the data is valid, create a range of the preceding 3 cells
         strRng = Range(Cells(.Row, .Column - 3), _
            Cells(.Row, .Column - 1)).Address
         intLen = Len(strRng)
         
         'Call MyFunction as a formula, with the range as the parameter.
         .Formula = "=MyFunction(" &; strRng &; ")"
            Application.SendKeys "{ENTER}"
      End If
   End With
End Sub
```

 **Value Entry:** This menu option is assigned the macro "Cbm_Active_Value", which enters the value produced by the UDF named "MyFunction" into the active cell. You must have values in the range B6:D6 and select cell E6 before clicking this menu option.



```vb
Sub Cbm_Active_Value()
   'Set up the variables.
   Dim intLen As Integer, strRng As String
   
   'Using the active cell, E6.
   With ActiveCell
      'If there isn't enough room in the column, then send a warning.
      If .Column < 4 Then
         Beep
         MsgBox "The function can be used only starting from column D!"
      
      'Otherwise, call MyFunction, using the range of the previous 3 cells.
      Else
         ActiveCell.Value = MyFunction(Range(ActiveCell.Offset(0, -3), _
            ActiveCell.Offset(0, -1)))
      End If
   End With
End Sub
```

 **Formula Selection:** This menu option is assigned the macro "Cbm_Formula_Select", which uses an InputBox for the user to select the range which the UDF "MyFunction" should calculate. The return value of the UDF is stored in the active cell.



```vb
Sub Cbm_Formula_Select()
   'Set up the variables.
   Dim rng As Range
   
   'Use the InputBox dialog to set the range for MyFunction, with some simple error handling.
   Set rng = Application.InputBox("Range:", Type:=8)
   If rng.Cells.Count <> 3 Then
      MsgBox "Length, width and height are needed -" &; _
         vbLf &; "please select three cells!"
      Exit Sub
   End If
   'Call MyFunction in the active cell, E6.
   ActiveCell.Formula = "=MyFunction(" &; rng.Address &; ")"
End Sub
```

 **Value Selection:** This menu option is assigned the macro "Cbm_Value_Select", which uses an InputBox for the user select the range which the UDF "MyFunction" should calculate. The value is stored in the active cell directly, instead of being returned by the UDF.



```vb
Sub Cbm_Value_Select()
   'Set up the variables.
   Dim rng As Range
   
   'Use the InputBox dialog to set the range for MyFunction, with some simple error handling.
   Set rng = Application.InputBox("Range:", Type:=8)
   If rng.Cells.Count <> 3 Then
     MsgBox "Length, width and height are needed -" &; _
         vbLf &; "please select three cells!"
      Exit Sub
   End If
   
   'Call MyFunction by value using the active cell, E6.
   ActiveCell.Value = MyFunction(rng)
End Sub
```


## About the Contributor
<a name="AboutContributor"> </a>

Holy Macro! Books publishes entertaining books for people who use Microsoft Office. See the complete catalog at MrExcel.com. 


