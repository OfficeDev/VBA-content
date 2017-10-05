---
title: Application.Replace Method (Project)
keywords: vbapj.chm241
f1_keywords:
- vbapj.chm241
ms.prod: project-server
api_name:
- Project.Application.Replace
ms.assetid: fd1c66ba-c611-ec97-ebb9-92ff0739c719
ms.date: 06/08/2017
---


# Application.Replace Method (Project)

Searches for an unfiltered value and replaces it with the specified value.


## Syntax

_expression_. **Replace** (**_Field_**, **_Test_**, **_Value_**, **_Replacement_**, **_ReplaceAll_**, **_Next_**, **_MatchCase_**, **_FieldID_**, **_TestID_**)

_expression_ A variable that represents an **Application** object.


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

Using the **Replace** method with no arguments, or without specifying _Field_, _Test_, and _Value_, displays the **Replace** dialog box that has options set for the previous state.

To replace a value in all available fields, use the **[ReplaceEx](application-replaceex-method-project.md)** method.


## Example

Either statement in the following example lowers the priority of all tasks that are equal to or over 800 to priority 600.

```vb
Sub LowerPriority() 
    Replace Field:="Priority", Test:="is greater than or equal to", Value:="800", _ 
        Replacement:="600", ReplaceAll:=True 
    Replace Field:="xx", Test:="xx", FieldID:=pjTaskPriority, TestID:=pjCompareGreaterThanOrEqual, _ 
        Value:="800", Replacement:="600" 
End Sub
```
