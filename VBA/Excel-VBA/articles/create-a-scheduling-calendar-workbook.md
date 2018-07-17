---
title: Create a Scheduling Calendar Workbook
ms.prod: excel
ms.assetid: 0f0f4946-c04c-4866-a6dd-79101df7bafb
ms.date: 06/08/2017
---


# Create a Scheduling Calendar Workbook

The following code example shows how to use information in one workbook to create a scheduling calendar workbook that contains one month per worksheet and can optionally include holidays and weekends.

 **Sample code provided by:** Holy Macro! Books, [Holy Macro! It's 2,500 Excel VBA Examples](http://www.mrexcel.com/store/index.php?l=product_detail&;p=1)

To run this code, your workbook must have a worksheet named "Cover" that contains the following:


- A spin control that contains a list of years name "SpinButton1"
    
- An option button for the "with weekends" option named "OptionButton1"
    
- An option button for the "without weekends" option named "OptionButton2"
    
- An option button for the "with holidays" option named "OptionButton3"
    
- An option button for the "without holidays" option named "OptionButton4"
    
Your workbook must also contain a worksheet named "Employee" that lists the names of the employees you want on your calendar in column A starting in cell A3, and a worksheet named "Holidays" that lists the dates of the holidays in column A starting in cell A2 and the name of the holidays in column B starting in cell B2.



```vb
Sub CreateCalendar()
   'Define your variables
   Dim wks As Worksheet
   Dim var As Variant
   Dim datDay As Date
   Dim iMonth As Integer, iCol As Integer, iCounter As Integer, iYear As Integer
   Dim sMonth As String
   Dim bln As Boolean
   
   'In the current application, turn off screen updating, save the current state of the status bar,
   'and then turn on the status bar.
   With Application
      .ScreenUpdating = False
      bln = .DisplayStatusBar
      .DisplayStatusBar = True
   End With
   
   'Initialize iYear with the value entered in the first spin button on the worksheet.
   iYear = Cover.SpinButton1.Value
   
   'Create a new workbook to hold your new calendar.
   Workbooks.Add
   
   'In this new workbook, clear out all the worksheets except for one.
   Application.DisplayAlerts = False
   For iCounter = 1 To Worksheets.Count - 1
      Worksheets(2).Delete
   Next iCounter
   Application.DisplayAlerts = True
   
   
   Set wks = ThisWorkbook.Worksheets("Employee")
   
   'For each month of the year
   For iMonth = 1 To 12
      'Create a new worksheet and label the worksheet tab with the name of the new month
      sMonth = Format(DateSerial(1, iMonth, 1), "mmmm")
      Application.StatusBar = "Place month " &; sMonth &; " on..."
      Worksheets.Add after:=Worksheets(Worksheets.Count)
      ActiveSheet.Name = sMonth
      
      'Copy the employee names to the first column, and add the dates across the remaining columns.
      wks.Range(wks.Cells(3, 1), wks.Cells( _
         WorksheetFunction.CountA(wks.Columns(1)) + 1, 1)).Copy Range("A2")
      Range("A1").Value = "'" &; ActiveSheet.Name &; " " &; iYear
      
      'Call the private subs, depending on what options are chosen for the calendar.
      
      'With weekends and holidays
      If Cover.OptionButton1.Value And Cover.OptionButton3.Value Then
         Call WithHW(iMonth)
      'With weekends, but without holidays
      ElseIf Cover.OptionButton1.Value And Cover.OptionButton3.Value = False Then
         Call WithWsansH(iMonth)
      'With holidays, but without weekends
      ElseIf Cover.OptionButton1.Value = False And Cover.OptionButton3.Value Then
         Call WithHsansW(iMonth)
      'Without weekends or holidays.
      Else
         Call SansWH(iMonth)
      End If
      
      'Apply some formatting.
      Rows(2).Value = Rows(1).Value
      Rows(2).NumberFormat = "ddd"
      Range("A2").Value = "Weekdays"
      Rows("1:2").Font.Bold = True
      Columns.AutoFit
   Next iMonth
   
   'Delete the first worksheet, because there was not anything in it.
   Application.DisplayAlerts = False
   Worksheets(1).Delete
   Application.DisplayAlerts = True
   
   'Label the window.
   Worksheets(1).Select
   ActiveWindow.Caption = "Yearly calendar " &; iYear
   
   'Do some final cleanup, and then close out the sub.
   With Application
      .ScreenUpdating = True
      .DisplayStatusBar = bln
      .StatusBar = False
   End With
End Sub


'Name: WithWH (with weekends and holidays)
'Description: Creates a calendar for the specified month, including both weekends and holidays.
Private Sub WithHW(ByVal iMonth As Integer)
   'Define your variables.
   Dim cmt As Comment
   Dim rng As Range
   Dim var As Variant
   Dim datDay As Date
   Dim iYear As Integer, iCol As Integer
   iCol = 1
   iYear = Cover.SpinButton1.Value
   
   'Go through every day of the month and put the date on the calendar in the first row.
   For datDay = DateSerial(iYear, iMonth, 1) To DateSerial(iYear, iMonth + 1, 0)
      iCol = iCol + 1
      Set rng = Range(Cells(1, iCol), Cells(WorksheetFunction.CountA(Columns(1)), iCol))
      
      'Determine if the day is a holiday.
      var = Application.Match(CDbl(datDay), ThisWorkbook.Worksheets("Holidays").Columns(1), 0)
      Cells(1, iCol).Value = datDay
      
      'Add the appropriate formatting that indicates a holiday or weekend.
      With rng.Interior
         Select Case Weekday(datDay)
            Case 1
               .ColorIndex = 35
            Case 7
               .ColorIndex = 36
         End Select
         If Not IsError(var) Then
            .ColorIndex = 34
            Set cmt = Cells(1, iCol).AddComment( _
               ThisWorkbook.Worksheets("Holidays").Cells(var, 2).Value)
            cmt.Shape.TextFrame.AutoSize = True
         End If
      End With
   Next datDay
End Sub


'Name: WithHsansW (with holidays, without weekends)
'Description: Creates a calendar for the specified month, including holidays, but not weekends.
Private Sub WithHsansW(ByVal iMonth As Integer)
   'Declare your variables.
   Dim datDay As Date
   Dim iYear As Integer, iCol As Integer
   iCol = 1
   iYear = Cover.SpinButton1.Value
   
   'For every day in the month, determine if the day is a weekend.
   For datDay = DateSerial(iYear, iMonth, 1) To DateSerial(iYear, iMonth + 1, 0)
      
      'If the day is not a weekend, put it on the calendar.
      If WorksheetFunction.Weekday(datDay, 2) < 6 Then
         iCol = iCol + 1
         Cells(1, iCol).Value = datDay
      End If
   Next datDay
End Sub


'Name: WithWsansH (with weekends, without holidays)
'Description: Creates a calendar for the specified month, including weekends, but not holidays.
Private Sub WithWsansH(ByVal iMonth As Integer)
   'Declare your variables.
   Dim var As Variant
   Dim datDay As Date
   Dim iYear As Integer, iCol As Integer
   iCol = 1
   iYear = Cover.SpinButton1.Value
   
   'For every day in the month, determine if the day is a holiday.
   For datDay = DateSerial(iYear, iMonth, 1) To DateSerial(iYear, iMonth + 1, 0)
      var = Application.Match(CDbl(datDay), ThisWorkbook.Worksheets("Holidays").Columns(1), 0)
      
      'If the day is not a holiday, put it on the calendar.
      If IsError(var) Then
         iCol = iCol + 1
         Cells(1, iCol).Value = datDay
      End If
   Next datDay
End Sub


'Name: SansWH (without weekends or holidays)
'Description: Creates a calendar for the specified month, not including weekends or holidays.
Private Sub SansWH(ByVal iMonth As Integer)
   'Set up your variables
   Dim var As Variant
   Dim datDay As Date
   Dim iYear As Integer, iCol As Integer
   iCol = 1
   iYear = Cover.SpinButton1.Value
   
   'For every day in the month, determine if the day is a weekend or a holiday.
   For datDay = DateSerial(iYear, iMonth, 1) To DateSerial(iYear, iMonth + 1, 0)
      If WorksheetFunction.Weekday(datDay, 2) < 6 Then
         var = Application.Match(CDbl(datDay), ThisWorkbook.Worksheets("Holidays").Columns(1), 0)
         
         'If the day is not a weekend or a holiday, put it on the calender.
         If IsError(var) Then
            iCol = iCol + 1
            Cells(1, iCol).Value = datDay
         End If
      End If
   Next datDay
End Sub
```


## About the Contributor
<a name="AboutContributor"> </a>

Holy Macro! Books publishes entertaining books for people who use Microsoft Office. See the complete catalog at MrExcel.com. 


