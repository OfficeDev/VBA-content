---
title: Application.FindEx Method (Project)
keywords: vbapj.chm97
f1_keywords:
- vbapj.chm97
ms.prod: project-server
api_name:
- Project.Application.FindEx
ms.assetid: fdb2661e-f705-ffa4-1ca3-7bbc97b9958d
ms.date: 06/08/2017
---


# Application.FindEx Method (Project)

Searches for an unfiltered value in a specified field or in all available fields; returns  **True** if the value is found.


## Syntax

_expression_. **FindEx** (**_Field_**, **_Test_**, **_Value_**, **_Next_**, **_MatchCase_**, **_FieldID_**, **_TestID_**, **_SearchAllFields_**)

_expression_ An expression that returns an **Application** object.


### Parameters

|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Field_|Optional|**String**|The name of the field to search.|
| _Test_|Optional|**String**|The type of comparison made between Field and Value. Can be one of the following [comparison strings](#comparison-strings).|
| _Value_|Optional|**String**|The value to compare with the field specified by Field.|
| _Next_|Optional|**Boolean**|**True** if Project searches down for the next occurrence of a value that matches the search criteria. **False** if Project searches up for the next occurrence. The default value is **True**.|
| _MatchCase_|Optional|**Boolean**|**True** if the search is case-sensitive. The default value is **False**.|
| _FieldID_|Optional|**Variant**|The field identification number can be one of the **[PjField](pjfield-enumeration-project.md)** constants. FieldID takes precedence over any Field value.|
| _TestID_|Optional|**Variant**|The test identification number can be one of the **[PjComparison](pjcomparison-enumeration-project.md)** constants. TestID takes precedence over any Test value.|
| _SearchAllFields_|Optional|**Boolean**|If **True**, search for the specified value in all available fields. The default value is **False**.|

<br/>

#### Comparison strings

|**Comparison string**|**Description**|
|:-----|:-----|
|"equals"|The value of _Field_ equals _Value_.|
|"does not equal"|The value of _Field_ does not equal _Value_.|
|"is greater than"|The value of _Field_ is greater than _Value_.|
|"is greater than or equal to"|The value of _Field_ is greater than or equal to _Value_.|
|"is less than"|The value of _Field_ is less than _Value_.|
|"is less than or equal to"|The value of _Field_ is less than or equal to _Value_.|
|"is within"|The value of _Field_ is within _Value_.|
|"is not within"|The value of _Field_ is not within _Value_.|
|"contains"|_Field_ contains _Value_.|
|"does not contain"|_Field_ does not contain _Value_.|
|"contains exactly"|_Field_ contains exactly _Value_.|


### Return value

 **Boolean**

## Remarks

Using the **FindEx** method with no arguments, or without specifying Field, Test, and Value, displays the **Find** dialog box that has options set for the previous state. If you set SearchAllFields to **True**, programmatic use still requires values for the Field, Test, and Value parameters.

## Example

Either of the following statements finds the next field that contains the value 2, within the set of all available fields.


```
FindEx Field:="Name", value:="2", Test:="contains", SearchAllFields:=True 
FindEx Field:="OtherField", value:="2", Test:="xx", FieldID:=pjTaskName, _
    TestID:=pjCompareContains, SearchAllFields:=True
```

