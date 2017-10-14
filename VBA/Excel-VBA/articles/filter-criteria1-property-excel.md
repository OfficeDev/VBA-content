---
title: Filter.Criteria1 Property (Excel)
keywords: vbaxl10.chm542074
f1_keywords:
- vbaxl10.chm542074
ms.prod: excel
api_name:
- Excel.Filter.Criteria1
ms.assetid: c1414fe3-92fd-e5cd-c60b-64e00cdf4973
ms.date: 06/08/2017
---


# Filter.Criteria1 Property (Excel)

Returns the first filtered value for the specified column in a filtered range. Read-only  **Variant** .


## Syntax

 _expression_ . **Criteria1**

 _expression_ A variable that represents a **Filter** object.


## Example

The following example sets a variable to the value of the  **Criteria1** property of the filter for the first column in the filtered range on the Crew worksheet.


```vb
With Worksheets("Crew") 
 If .AutoFilterMode Then 
 With .AutoFilter.Filters(1) 
 If .On Then c1 = .Criteria1 
 End With 
 End If 
End With
```


## See also


#### Concepts


[Filter Object](filter-object-excel.md)

