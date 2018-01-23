---
title: Application.SetAutoFilter Method (Project)
keywords: vbapj.chm2166
f1_keywords:
- vbapj.chm2166
ms.prod: project-server
api_name:
- Project.Application.SetAutoFilter
ms.assetid: 4e4b4d4a-838b-f9b7-e3ab-d7bfa8efce5f
ms.date: 06/08/2017
---


# Application.SetAutoFilter Method (Project)

Sets the criteria for an AutoFilter for a specified field in a sheet view.

## Syntax

_expression_. **SetAutoFilter** (**_FieldName_**, **_FilterType_**, **_Test1_**, **_Criteria1_**, **_Operation_**, **_Test2_**, **_Criteria2_**)

_expression_ An expression that returns an **Application** object.


### Parameters

|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FieldName_|Required|**String**|Name of the field.|
| _FilterType_|Optional|**PjAutoFilterType**|Type of filter; can be one of the **[PjAutoFilterType](pjautofiltertype-enumeration-project.md)** constants. The default value is **pjAutoFilterClear**, which clears the AutoFilter.|
| _Test1_|Optional|**String**|Specifies the type of comparison for the first test. Requires that _FilterType_ is **pjAutoFilterCustom**, and that _Criteria1_ specifies a value. Can be one of the [comparison strings](#comparison-strings).|
| _Criteria1_|Optional|**String**|The value of the first comparison with the value of the field specified by _FieldName_.|
| _Operation_|Optional|**String**|The logical operation if there is a second test. The _Operation_ value can be "And" or "Or".|
| _Test2_|Optional|**String**|Specifies the type of comparison for the second test. Requires that _FilterType_ is **pjAutoFilterCustom**, the _Operation_ value must be set, and that _Criteria2_ specifies a value. The string can be one of the comparisons in the table for _Test1_.|
| _Criteria2_|Optional|**String**|The value of the second comparison with the value of the field specified by _FieldName_.|

<br/>

#### Comparison strings

|**Comparison string**|**Description**|
|:-----|:-----|
|"equals"|The value of _FieldName_ equals _Criteria1_.|
|"does not equal"|The value of _FieldName_ does not equal _Criteria1_.|
|"is greater than"|The value of _FieldName_ is greater than _Criteria1_.|
|"is greater than or equal to"|The value of _FieldName_ is greater than or equal to _Criteria1_.|
|"is less than"|The value of _FieldName_ is less than _Criteria1_.|
|"is less than or equal to"|The value of _FieldName_ is less than or equal to _Criteria1_.|
|"is within"|The value of _FieldName_ is within _Criteria1_.|
|"is not within"|The value of _FieldName_ is not within _Criteria1_.|

<br/>

### Return value

 **Boolean**


## Remarks

To turn the AutoFilter feature on or off, see the **[AutoFilter](application-autofilter-method-project.md)** method.

> [!NOTE]
> A column name in a sheet view can have a different title than the name of the field it shows.


## Example

The following example sets a custom AutoFilter for the "% Work Complete" field. 

```vb
Sub TestAutoFilter() 
    If Not ActiveProject.AutoFilter Then 
        Application.AutoFilter 
    End If 
 
    Application.SetAutoFilter FieldName:="% Work Complete", FilterType:=pjAutoFilterCustom, _ 
    Test1:="equals", Criteria1:="0%" 
End Sub
```

If there is an AutoFilter set for the "% Work Complete" field, the following line of code clears the AutoFilter because the default value for the optional _FilterType_ argument is **pjAutoFilterClear**.

```vb
Application.SetAutoFilter FieldName:="% Work Complete"
```


