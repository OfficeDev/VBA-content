---
title: Filters Object (Excel)
keywords: vbaxl10.chm539072
f1_keywords:
- vbaxl10.chm539072
ms.prod: excel
api_name:
- Excel.Filters
ms.assetid: a714ed69-7772-5ade-3acd-f3e3d98db62c
ms.date: 06/08/2017
---


# Filters Object (Excel)

A collection of  **[Filter](filter-object-excel.md)** objects that represents all the filters in an autofiltered range.


## Example

Use the  **[Filters](autofilter-filters-property-excel.md)** property to return the **Filters** collection. The following example creates a list that contains the criteria and operators for the filters in the autofiltered range on the Crew worksheet.


```
Dim f As Filter 
Dim w As Worksheet 
Const ns As String = "Not set" 
 
Set w = Worksheets("Crew") 
Set w2 = Worksheets("FilterData") 
rw = 1 
For Each f In w.AutoFilter.Filters 
 If f.On Then 
 c1 = Right(f.Criteria1, Len(f.Criteria1) - 1) 
 If f.Operator Then 
 op = f.Operator 
 c2 = Right(f.Criteria2, Len(f.Criteria2) - 1) 
 Else 
 op = ns 
 c2 = ns 
 End If 
 Else 
 c1 = ns 
 op = ns 
 c2 = ns 
 End If 
 w2.Cells(rw, 1) = c1 
 w2.Cells(rw, 2) = op 
 w2.Cells(rw, 3) = c2 
 rw = rw + 1 
Next
```

Use  **Filters** ( _index_ ), where _index_ is the filter title or index number, to return a single **Filter** object. The following example sets a variable to the value of the **On** property of the filter for the first column in the filtered range on the Crew worksheet.




```
Set w = Worksheets("Crew") 
If w.AutoFilterMode Then 
 filterIsOn = w.AutoFilter.Filters(1).On 
End If
```


## Properties



|**Name**|
|:-----|
|[Application](filters-application-property-excel.md)|
|[Count](filters-count-property-excel.md)|
|[Creator](filters-creator-property-excel.md)|
|[Item](filters-item-property-excel.md)|
|[Parent](filters-parent-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
