---
title: Application.FilterEdit Method (Project)
keywords: vbapj.chm503
f1_keywords:
- vbapj.chm503
ms.prod: project-server
api_name:
- Project.Application.FilterEdit
ms.assetid: e576d3e2-5ac9-006a-2151-dc918b71eef8
ms.date: 06/08/2017
---


# Application.FilterEdit Method (Project)

Creates, edits, or copies a filter.

## Syntax

_expression_. **FilterEdit** (**_Name_**, **_TaskFilter_**, **_Create_**, **_OverwriteExisting_**, **_Parenthesis_**, **_NewName_**, **_FieldName_**, **_NewFieldName_**, **_Test_**, **_Value_**, **_Operation_**, **_ShowInMenu_**, **_ShowSummaryTasks_**)

_expression_ A variable that represents an **Application** object.


### Parameters

|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|The name of a filter to edit, create, or copy.|
| _TaskFilter_|Required|**Boolean**|**True** if the filter specified with Name contains task information. **False** if the filter contains resource information.|
| _Create_|Optional|**Boolean**|**True** if a new filter is created. The new filter is a copy of the filter specified with Name and is given the name specified with NewName. If NewName is empty, the new filter is given the name specified by Name. The default value is **False**.|
| _OverwriteExisting_|Optional|**Boolean**|**True** if the existing filter is overwritten with a new filter. The default value is **False**.|
| _Parenthesis_|Optional|**Boolean**|**True** if the criterion established by FieldName, Test, and Value is evaluated as a parenthetical **AND** or **OR** clause (the value specified with Operation) in relation to other criteria, in the manner of (a AND b) OR c.|
| _NewName_|Optional|**String**| A new name for the filter specified with Name (Create is **False** ) or a name for the new filter (Create is **True** ). If NewName is empty and Create is **False**, the filter specified with Name retains its current name. The default value is **Empty**.|
| _FieldName_|Optional|**String**|The name of a field to change.|
| _NewFieldName_|Optional|**String**| A new name for the field specified with FieldName.|
| _Test_|Required|**String**| The type of comparison made between FieldName and Value that acts as selection criteria for the filter. Can be one of the [comparison strings](#comparison-strings).|
| _Value_|Optional|**String**| The value to compare with the value of the field specified with FieldName.|
| _Operation_|Optional|**String**| How the criterion established with FieldName, Test, and Value relates to other criteria in the filter. The Operation argument can be set to "And" or "Or".|
| _ShowInMenu_|Optional|**Boolean**|**True** if the filter is displayed in the **Filter** drop-down list. The default value is **False**. **Note** To display the list of filters, on the Ribbon, on the **View** tab, click the **Filter** drop-down list.|
| _ShowSummaryTasks_|Optional|**Boolean**|**True** if the summary tasks of the filtered tasks are displayed. The default value is **False**.|

<br/>

#### Comparison strings

|**Comparison string**|**Description**|
|:-----|:-----|
|"equals"|The value of _FieldName_ equals _Value_.|
|"does not equal"|The value of _FieldName_ does not equal _Value_.|
|"is greater than"|The value of _FieldName_ is greater than _Value_.|
|"is greater than or equal to"|The value of _FieldName_ is greater than or equal to _Value_.|
|"is less than"|The value of _FieldName_ is less than _Value_.|
|"is less than or equal to"|The value of _FieldName_ is less than or equal to _Value_.|
|"is within"|The value of _FieldName_ is within _Value_.|
|"is not within"|The value of _FieldName_ is not within _Value_.|
|"contains"|_FieldName_ contains _Value_.|
|"does not contain"|_FieldName_ does not contain _Value_.|
|"contains exactly"|_FieldName_ exactly contains _Value_.|


### Return value

 **Boolean**


## Example

The following example creates a filter (if one does not exist) for tasks with the highest priority, and then applies the filter.


```vb 
Sub CreateAndApplyHighestPriorityFilter() 
    Dim TaskFilter As Variant  ' Index for For Each loop. 
    Dim Found As Boolean    ' Whether or not the filter exists. 
    Found = False   ' Assume the filter does not exist. 
    ' Look for filter. 
    For Each TaskFilter In ActiveProject.TaskFilterList 
        If TaskFilter = "Highest Priority" Then 
            Found = True 
            Exit For 
        End If 
    Next TaskFilter 
 
    ' If filter doesn't exist, create it. 
    If Not Found Then FilterEdit Name:="Highest Priority", _ 
        Create:=True, TaskFilter:=True, FieldName:="Priority", _ 
        Test:="equals", Value:="Highest" 
    FilterApply "Highest Priority" 
End Sub    
```


