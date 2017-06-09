---
title: Range.Group Method (Excel)
keywords: vbaxl10.chm144141
f1_keywords:
- vbaxl10.chm144141
ms.prod: excel
api_name:
- Excel.Range.Group
ms.assetid: da736f64-35df-ecaf-88ac-8c61f7d3c0d0
ms.date: 06/08/2017
---


# Range.Group Method (Excel)

When the  **[Range](range-object-excel.md)** object represents a single cell in a PivotTable field?s data range, the **Group** method performs numeric or date-based grouping in that field.


## Syntax

 _expression_ . **Group**( **_Start_** , **_End_** , **_By_** , **_Periods_** )

 _expression_ A variable that represents a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Start_|Optional| **Variant**|The first value to be grouped. If this argument is omitted or  **True** , the first value in the field is used.|
| _End_|Optional| **Variant**|The last value to be grouped. If this argument is omitted or  **True** , the last value in the field is used.|
| _By_|Optional| **Variant**|If the field is numeric, this argument specifies the size of each group. If the field is a date, this argument specifies the number of days in each group if element 4 in the  _Periods_ array is **True** and all the other elements are **False** . Otherwise, this argument is ignored. If this argument is omitted, Microsoft Excel automatically chooses a default group size.|
| _Periods_|Optional| **Variant**|An array of  **Boolean** values that specify the period for the group, described in the Remarks section. If an element in the array is **True** , a group is created for the corresponding time; if the element is **False** , no group is created. If the field isn?t a date field, this argument is ignored.|

### Return Value

Variant


## Remarks

The  **Boolean** array for the _Periods_ parameter contains the following elements:



|**Array element**|**Period**|
|:-----|:-----|
|1|Seconds|
|2|Minutes|
|3|Hours|
|4|Days|
|5|Months|
|6|Quarters|
|7|Years|
Because a group of shapes is treated as a single shape, grouping and ungrouping shapes changes the number of items in the  **[Shapes](shapes-object-excel.md)** collection and changes the index numbers of items that come after the affected items in the collection.

The  **Range** object must be a single cell in the PivotTable field?s data range. If you attempt to apply this method to more than one cell, it will fail (without displaying an error message).


## See also


#### Concepts


[Range Object](range-object-excel.md)

