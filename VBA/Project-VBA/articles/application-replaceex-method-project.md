---
title: Application.ReplaceEx Method (Project)
keywords: vbapj.chm98
f1_keywords:
- vbapj.chm98
ms.prod: project-server
api_name:
- Project.Application.ReplaceEx
ms.assetid: af284688-0701-abc7-4d04-b258957fa9dc
ms.date: 06/08/2017
---


# Application.ReplaceEx Method (Project)

Searches for an unfiltered value in a specified field, or in all available fields, and replaces it with the specified value.

## Syntax

_expression_. **ReplaceEx** (**_Field_**, **_Test_**, **_Value_**, **_Replacement_**, **_ReplaceAll_**, **_Next_**, **_MatchCase_**, **_FieldID_**, **_TestID_**, **_SearchAllFields_**)

_expression_ An expression that returns an **Application** object.


### Parameters

|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Field_|Optional|**String**|The name of the field to search.|
| _Test_|Optional|**String**|The type of comparison made between _Field_ and _Value_. Can be one of the [comparison strings](#comparison-strings).|
| _Value_|Optional|**Variant**|The value to compare with the value of the field specified in _Field_.|
| _Replacement_|Optional|**Variant**|Use "" (an empty string) to clear _Field_ where it meets the test specified by _Test_ and _Value_.|
| _ReplaceAll_|Optional|**Variant**|**True** if all occurrences of _Value_ are replaced. **False** if only the first occurrence is replaced. The default value is **False**.|
| _Next_|Optional|**Variant**|**True** if Project searches down for the next occurrence of matching search criteria. **False** if Project searches up for the next occurrence. The default value is **True**.|
| _MatchCase_|Optional|**Variant**|**True** if the search is case-sensitive. The default value is **False**.|
| _FieldID_|Optional|**Variant**|The field identification number can be one of the **[PjField](pjfield-enumeration-project.md)** constants. _FieldID_ takes precedence over any _Field_ value.|
| _TestID_|Optional|**Variant**|The test identification number can be one of the **[PjComparison](pjcomparison-enumeration-project.md)** constants. _TestID_ takes precedence over any _Test_ value.|
| _SearchAllFields_|Optional|**Variant**|If **True**, replace the specified value in all available fields. The default value is **False**. _SearchAllFields_ takes precedence over _Field_ and _FieldID_.|

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

<br/>

### Return value

 **Boolean**

 **True** if any replacements are made; otherwise, **False**.


## Remarks

Using the **ReplaceEx** method with no arguments, or without specifying _Field_, _Test_, and _Value_, displays the **Replace** dialog box that has options set for the previous state. If you set _SearchAllFields_ to **True**, programmatic use still requires values for the _Field_, _Test_, and _Value_ parameters.


## Example

Either line in the following example replaces "Bad" with "Good", within the set of all available fields.


```vb
Sub Bad2Good() 
    ReplaceEx Field:="Name", Test:="contains", Value:="Bad", Replacement:="Good", _ 
        ReplaceAll:=True, SearchAllFields:=True 
    ReplaceEx Field:="xx", Test:="xx", TestID:=pjCompareContains, Value:="Bad", Replacement:="Good", _ 
        ReplaceAll:=True, SearchAllFields:=True 
End Sub
```


