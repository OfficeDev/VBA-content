---
title: Filter Object (Excel)
keywords: vbaxl10.chm541072
f1_keywords:
- vbaxl10.chm541072
ms.prod: excel
api_name:
- Excel.Filter
ms.assetid: 950023f9-a984-01fa-aa77-947cbbff0433
ms.date: 06/08/2017
---


# Filter Object (Excel)

Represents a filter for a single column.


## Remarks

 The **Filter** object is a member of the **[Filters](filters-object-excel.md)** collection. The **Filters** collection contains all the filters in an autofiltered range.


## Example

Use  **[Filters](autofilter-filters-property-excel.md)** ( _index_ ), where _index_ is the filter title or index number, to return a single **Filter** object. The following example sets a variable to the value of the **[On](filter-on-property-excel.md)** property of the filter for the first column in the filtered range on the Crew worksheet.


```
Set w = Worksheets("Crew") 
If w.AutoFilterMode Then 
 filterIsOn = w.AutoFilter.Filters(1).On 
End If
```

Note that all the properties of the  **Filter** object are read-only. To set these properties, apply autofiltering manually or using the **[AutoFilter](range-autofilter-method-excel.md)** method of the **[Range](range-object-excel.md)** object, as shown in the following example.




```
Set w = Worksheets("Crew") 
w.Cells.AutoFilter field:=2, Criteria1:="Crucial", _ 
 Operator:=xlOr, Criteria2:="Important"
```


## Properties



|**Name**|
|:-----|
|[Application](filter-application-property-excel.md)|
|[Count](filter-count-property-excel.md)|
|[Creator](filter-creator-property-excel.md)|
|[Criteria1](filter-criteria1-property-excel.md)|
|[Criteria2](filter-criteria2-property-excel.md)|
|[On](filter-on-property-excel.md)|
|[Operator](filter-operator-property-excel.md)|
|[Parent](filter-parent-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
