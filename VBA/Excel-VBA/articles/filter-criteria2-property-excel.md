---
title: Filter.Criteria2 Property (Excel)
keywords: vbaxl10.chm542076
f1_keywords:
- vbaxl10.chm542076
ms.prod: excel
api_name:
- Excel.Filter.Criteria2
ms.assetid: 73bd97f8-8ee7-b2a0-8f9c-6a20e3e11d09
ms.date: 06/08/2017
---


# Filter.Criteria2 Property (Excel)

Returns the second filtered value for the specified column in a filtered range. Read-only  **Variant** .


## Syntax

 _expression_ . **Criteria2**

 _expression_ A variable that represents a **Filter** object.


## Remarks

If you try to access the  **Criteria2** property for a filter that does not use two criteria, an error will occur. Check that the **[Operator](filter-operator-property-excel.md)** property of a **Filter** object doesn't equal zero (0) before trying to access the **Criteria2** property.


## Example

The following example sets a variable to the value of the  **Criteria2** property of the filter for the first column in the filtered range on the Crew worksheet.


```vb
With Worksheets("Crew") 
 If .AutoFilterMode Then 
 With .AutoFilter.Filters(1) 
 If .On And .Operator Then 
 c2 = .Criteria2 
 Else 
 c2 = "Not set" 
 End If 
 End With 
 End If 
End With
```


## See also


#### Concepts


[Filter Object](filter-object-excel.md)

