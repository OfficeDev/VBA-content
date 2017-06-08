---
title: PivotFilters.Add Method (Excel)
keywords: vbaxl10.chm772078
f1_keywords:
- vbaxl10.chm772078
ms.prod: excel
api_name:
- Excel.PivotFilters.Add
ms.assetid: bf3bb727-4c00-1f8e-5acd-af0b974cba5b
ms.date: 06/08/2017
---


# PivotFilters.Add Method (Excel)

Adds new filters to the  **PivotFilters** collection.


## Syntax

 _expression_ . **Add2**( **_Type_** , **_DataField_** , **_Value1_** , **_Value2_** , **_Order_** , **_Name_** , **_Description_** , **_MemberPropertyField_** , **_WholeDayFilter_** )

 _expression_ A variable that represents a **PivotFilters** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **XlPivotFilterType**|Requires an  **[XlPivotFilterType](xlpivotfiltertype-enumeration-excel.md)** type of filter.|
| _DataField_|Optional| **Variant**|The field to which the filter is attached.|
| _Value1_|Optional| **Variant**|Filter value 1.|
| _Value2_|Optional| **Variant**|Filter value 2.|
| _Order_|Optional| **Variant**|Order in which the data should be filtered.|
| _Name_|Optional| **Variant**|Name of the filter.|
| _Description_|Optional| **Variant**|A brief description of the filter.|
| _MemberPropertyField_|Optional| **Variant**|Specifies the member property field on which the label filter is based.|
| _WholeDayFilter_|Optional| **Variant**|Specifies a filter based on days.|

### Return Value

PivotFilter


## Example

Following are some examples of how to use the  **Add** function correctly.


```vb
ActiveCell.PivotField.PivotFilters.Add FilterType := xlThisWeek 
 
ActiveCell.PivotField.PivotFilters.Add FilterType := xlTopCount DataField := MyPivotField2 Value1 := 10 
 
ActiveCell.PivotField.PivotFilters.Add FilterType := xlCaptionIsNotBetween Value1 := "A" Value2 := "G" 
 
ActiveCell.PivotField.PivotFilters.Add FilterType := xlValueIsGreaterThanOrEqualTo DataField := MyPivotField2 Value1 := 10000 

```

The following example returns a run-time error because the data type of  _Value1_ is invalid.




```vb
ActiveCell.PivotField.PivotFilters.Add FilterType := xlValueIsGreaterThanOrEqualTo DataField := MyPivotField2 Value1 := ?Allan?
```


## See also


#### Concepts


[PivotFilters Object](pivotfilters-object-excel.md)

