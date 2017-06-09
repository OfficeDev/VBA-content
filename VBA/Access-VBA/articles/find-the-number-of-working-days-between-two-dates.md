---
title: Find the Number of Working Days Between Two Dates
ms.prod: access
ms.assetid: 2831d409-1b10-06ef-54ec-9c3386e70021
ms.date: 06/08/2017
---


# Find the Number of Working Days Between Two Dates

Access does not have a built-in function to determine the number of working days between two dates. The following user-defined function illustrates how to calculate the number of working days between two dates. 


 **Note**  This function does not account for holidays.


```vb
Function Work_Days(BegDate As Variant, EndDate As Variant) As Integer 
 
 Dim WholeWeeks As Variant 
 Dim DateCnt As Variant 
 Dim EndDays As Integer 
 
 On Error GoTo Err_Work_Days 
 
 BegDate = DateValue(BegDate) 
 EndDate = DateValue(EndDate) 
 WholeWeeks = DateDiff("w", BegDate, EndDate) 
 DateCnt = DateAdd("ww", WholeWeeks, BegDate) 
 EndDays = 0 
 
 Do While DateCnt <= EndDate 
 If Format(DateCnt, "ddd") <> "Sun" And _ 
 Format(DateCnt, "ddd") <> "Sat" Then 
 EndDays = EndDays + 1 
 End If 
 DateCnt = DateAdd("d", 1, DateCnt) 
 Loop 
 
 Work_Days = WholeWeeks * 5 + EndDays 
 
Exit Function 
 
 Err_Work_Days: 
 
 ' If either BegDate or EndDate is Null, return a zero 
 ' to indicate that no workdays passed between the two dates. 
 
 If Err.Number = 94 Then 
 Work_Days = 0 
 Exit Function 
 Else 
' If some other error occurs, provide a message. 
 MsgBox "Error " &; Err.Number &; ": " &; Err.Description 
 End If 
 
End Function
```


