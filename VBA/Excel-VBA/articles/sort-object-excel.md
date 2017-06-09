---
title: Sort Object (Excel)
keywords: vbaxl10.chm846072
f1_keywords:
- vbaxl10.chm846072
ms.prod: excel
api_name:
- Excel.Sort
ms.assetid: 637ee681-743c-5196-2bfc-4a5bea025295
ms.date: 06/08/2017
---


# Sort Object (Excel)

Represents a sort of a range of data.


## Example

The following proceedure builds and sorts data in a range in the active worksheet.


```
Sub SortData() 
 
 'Building data to sort on the active sheet. 
 Range("A1").Value = "Name" 
 Range("A2").Value = "Bill" 
 Range("A3").Value = "Rod" 
 Range("A4").Value = "John" 
 Range("A5").Value = "Paddy" 
 Range("A6").Value = "Kelly" 
 Range("A7").Value = "William" 
 Range("A8").Value = "Janet" 
 Range("A9").Value = "Florence" 
 Range("A10").Value = "Albert" 
 Range("A11").Value = "Mary" 
 MsgBox "The list is out of order. Hit Ok to continue...", vbInformation 
 
 'Selecting a cell within the range. 
 Range("A2").Select 
 
 'Applying sort. 
 With ActiveWorkbook.Worksheets(ActiveSheet.Name).Sort 
 .SortFields.Clear 
 .SortFields.Add Key:=Range("A2:A11"), _ 
 SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal 
 .SetRange Range("A1:A11") 
 .Header = xlYes 
 .MatchCase = False 
 .Orientation = xlTopToBottom 
 .SortMethod = xlPinYin 
 .Apply 
 End With 
 MsgBox "Sort complete.", vbInformation 
 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Apply](sort-apply-method-excel.md)|
|[SetRange](sort-setrange-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](sort-application-property-excel.md)|
|[Creator](sort-creator-property-excel.md)|
|[Header](sort-header-property-excel.md)|
|[MatchCase](sort-matchcase-property-excel.md)|
|[Orientation](sort-orientation-property-excel.md)|
|[Parent](sort-parent-property-excel.md)|
|[Rng](sort-rng-property-excel.md)|
|[SortFields](sort-sortfields-property-excel.md)|
|[SortMethod](sort-sortmethod-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
