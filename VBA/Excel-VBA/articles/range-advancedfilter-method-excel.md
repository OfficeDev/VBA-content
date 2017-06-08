---
title: Range.AdvancedFilter Method (Excel)
keywords: vbaxl10.chm144078
f1_keywords:
- vbaxl10.chm144078
ms.prod: excel
api_name:
- Excel.Range.AdvancedFilter
ms.assetid: fe1a19fc-ab0f-6149-25d9-6102d5789757
ms.date: 06/08/2017
---


# Range.AdvancedFilter Method (Excel)

Filters or copies data from a list based on a criteria range. If the initial selection is a single cell, that cell's current region is used.


## Syntax

 _expression_ . **AdvancedFilter**( **_Action_** , **_CriteriaRange_** , **_CopyToRange_** , **_Unique_** )

 _expression_ A variable that represents a **Range** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Action_|Required| **[XlFilterAction](xlfilteraction-enumeration-excel.md)**|One of the constants of  **XlFilterAction** specifying whether to make a copy or filter the list in place.|
| _CriteriaRange_|Optional| **Variant**|The criteria range. If this argument is omitted, there are no criteria.|
| _CopyToRange_|Optional| **Variant**|The destination range for the copied rows if  _Action_ is **xlFilterCopy** . Otherwise, this argument is ignored.|
| _Unique_|Optional| **Variant**| **True** to filter unique records only. **False** to filter all records that meet the criteria. The default value is **False** .|

### Return Value

Variant


## Example

This example filters a database (named "Database") based on a criteria range named "Criteria."


```vb
Range("Database").AdvancedFilter _ 
 Action:=xlFilterInPlace, _ 
 CriteriaRange:=Range("Criteria")
```


## See also


#### Concepts


[Range Object](range-object-excel.md)

