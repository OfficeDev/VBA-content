---
title: Calculate Age
ms.prod: access
ms.assetid: 4afca7f2-9864-6300-79c4-c4e251b0b66d
ms.date: 06/08/2017
---


# Calculate Age

Access does not include a function that will calculate the age of a person or thing based on a given date. This topic contains Visual Basic for Applications (VBA) code for two custom functions,  **Age** and **AgeMonths**, that will calculate age based on a given date.

The following function calculates age in years from a given date to today's date.



```vb
 Function Age (varBirthDate As Variant) As Integer 
 Dim varAge As Variant 
 
 If IsNull(varBirthdate) then Age = 0: Exit Function 
 
 varAge = DateDiff("yyyy", varBirthDate, Now) 
 If Date < DateSerial(Year(Now), Month(varBirthDate), _ 
 Day(varBirthDate)) Then 
 varAge = varAge - 1 
 End If 
 Age = CInt(varAge) 
 End Function
```

The following function calculates the number of months that have transpired since the last month supplied by the given date. If the given date is a birthday, the function returns the number of months since the last birthday.



```vb
 Function AgeMonths(ByVal StartDate As String) As Integer 
 Dim tAge As Double 
 tAge = (DateDiff("m", StartDate, Now)) 
 If (DatePart("d", StartDate) > DatePart("d", Now)) Then 
 tAge = tAge - 1 
 End If 
 
 If tAge < 0 Then 
 tAge = tAge + 1 
 End If 
 
 AgeMonths = CInt(tAge Mod 12) 
 
 End Function
```


