---
title: AutoFilter Object (Excel)
keywords: vbaxl10.chm537072
f1_keywords:
- vbaxl10.chm537072
ms.prod: excel
api_name:
- Excel.AutoFilter
ms.assetid: 1a6fcf3b-52be-b599-029b-a3c53d12f85e
ms.date: 06/08/2017
---


# AutoFilter Object (Excel)

Represents autofiltering for the specified worksheet.


 **Note**  When using  **AutoFilter** with dates, the format should be consistent with English date separators ("/") instead of local settings ("."). A valid date would be "2/2/2007", whereas "2.2.2007" is invalid.


 **Note**  Working with objects (e g  **Interior** Object) requires adding a reference to an object. You will find more information about assigning an Object reference to a variable or property in the[Set Statement](http://msdn.microsoft.com/library/59de2927-b338-0038-50b9-3379d7331935%28Office.15%29.aspx).


## Example

Use the  **[AutoFilter](worksheet-autofilter-property-excel.md)** property to return the **AutoFilter** object. Use the **[Filters](autofilter-filters-property-excel.md)** property to return a collection of individual column filters. Use the **[Range](autofilter-range-property-excel.md)** property to return the **Range** object that represents the entire filtered range. The following example stores the address and filtering criteria for the current filtering and then applies new filters.


```
Dim w As Worksheet 
Dim filterArray() 
Dim currentFiltRange As String 
 
Sub ChangeFilters() 
 
Set w = Worksheets("Crew") 
With w.AutoFilter 
 currentFiltRange = .Range.Address 
 With .Filters 
 ReDim filterArray(1 To .Count, 1 To 3) 
 For f = 1 To .Count 
 With .Item(f) 
 If .On Then 
 filterArray(f, 1) = .Criteria1 
 If .Operator Then 
 filterArray(f, 2) = .Operator 
 filterArray(f, 3) = .Criteria2 
 End If 
 End If 
 End With 
 Next 
 End With 
End With 
 
w.AutoFilterMode = False 
w.Range("A1").AutoFilter field:=1, Criteria1:="S" 
 
End Sub
```

To create an  **AutoFilter** object for a worksheet, you must turn autofiltering on for a range on the worksheet either manually or using the **[AutoFilter](range-autofilter-method-excel.md)** method of the **[Range](range-object-excel.md)** object. The following example uses the values stored in module-level variables in the previous example to restore the original autofiltering to the Crew worksheet.




```
Sub RestoreFilters() 
Set w = Worksheets("Crew") 
w.AutoFilterMode = False 
For col = 1 To UBound(filterArray(), 1) 
 If Not IsEmpty(filterArray(col, 1)) Then 
 If filterArray(col, 2) Then 
 w.Range(currentFiltRange).AutoFilter field:=col, _ 
 Criteria1:=filterArray(col, 1), _ 
 Operator:=filterArray(col, 2), _ 
 Criteria2:=filterArray(col, 3) 
 Else 
 w.Range(currentFiltRange).AutoFilter field:=col, _ 
 Criteria1:=filterArray(col, 1) 
 End If 
 End If 
Next 
End Sub 

```


 **Note**  When using  **AutoFilter** with dates, the format should be consistent with English date separators ("/") instead of local settings ("."). A valid date would be "2/2/2007", whereas "2.2.2007" is invalid.


## Methods



|**Name**|
|:-----|
|[ApplyFilter](autofilter-applyfilter-method-excel.md)|
|[ShowAllData](autofilter-showalldata-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[Application](autofilter-application-property-excel.md)|
|[Creator](autofilter-creator-property-excel.md)|
|[FilterMode](autofilter-filtermode-property-excel.md)|
|[Filters](autofilter-filters-property-excel.md)|
|[Parent](autofilter-parent-property-excel.md)|
|[Range](autofilter-range-property-excel.md)|
|[Sort](autofilter-sort-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
