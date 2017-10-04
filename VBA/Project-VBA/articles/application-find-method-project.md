---
title: Application.Find Method (Project)
keywords: vbapj.chm215
f1_keywords:
- vbapj.chm215
ms.prod: project-server
api_name:
- Project.Application.Find
ms.assetid: 0e7b1027-5609-19fa-f100-4eb7b108bae7
ms.date: 06/08/2017
---


# Application.Find Method (Project)

Searches for an unfiltered value; returns  **True** if the value is found.


## Syntax

_expression_. **Find** (**_Field_**, **_Test_**, **_Value_**, **_Next_**, **_MatchCase_**, **_FieldID_**, **_TestID_**)

_expression_ A variable that represents an **Application** object.


### Parameters

|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Field_|Optional|**String**|The name of the field to search.|
| _Test_|Optional|**String**|The type of comparison made between Field and Value. Can be one of the [comparison strings](#comparison-strings).|
| _Value_|Optional|**String**|The value to compare with the field specified by Field.|
| _Next_|Optional|**Boolean**|**True** if Project searches down for the next occurrence of a value that matches the search criteria. **False** if Project searches up for the next occurrence. The default value is **True**.|
| _MatchCase_|Optional|**Boolean**|**True** if the search is case-sensitive. The default value is **False**.|
| _FieldID_|Optional|**Variant**|The field identification number can be one of the **[PjField](pjfield-enumeration-project.md)** constants. FieldID takes precedence over any Field value.|
| _TestID_|Optional|**Variant**|The test identification number can be one of the **[PjComparison](pjcomparison-enumeration-project.md)** constants. TestID takes precedence over any Test value.|

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

## Remarks

Using the **Find** method with no arguments, or without specifying Field, Test, and Value, displays the **Find** dialog box that has options set for the previous state.

To find a value where you can search all available fields, use the **[FindEx](application-findex-method-project.md)** method.

## Example

Either statement in the following example finds the next task with priority = 600.

```vb
Sub FindFieldByPriority 
 Find Field:="Priority", Test:="equals", Value:="600" 
 Find Field:="xx", Test:="xx", FieldID:=pjTaskPriority, TestID:=pjCompareEquals, Value:="600" 
End Sub
```


