---
title: Application.SetTPField Method (Project)
keywords: vbapj.chm1513
f1_keywords:
- vbapj.chm1513
ms.prod: project-server
api_name:
- Project.Application.SetTPField
ms.assetid: 66867c0a-e5a7-9492-463b-0cb955f020df
ms.date: 06/08/2017
---


# Application.SetTPField Method (Project)

Sets a value for the percent complete field of one or more tasks in the Team Planner view.


## Syntax

 _expression_. **SetTPField**( ** _Field_**, ** _Value_**, ** _AllSelectedTasks_**, ** _Create_**, ** _TaskID_**, ** _ProjectName_** )

 _expression_ An expression that returns an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Field_|Required|**String**|The  _Field_ argument can be only "% Complete". You can also use the **FieldConstantToFieldName** method with **pjTaskPercentComplete**, as in the example.|
| _Value_|Required|**String**|Value of the  _Field_ argument. Values can range from "0" to "100" for percent complete.|
| _AllSelectedTasks_|Optional|**Boolean**|Not used in Project. The value is  **True**, which means that the _Field_ and _Value_ arguments are set for all selected tasks.|
| _Create_|Optional|**Boolean**|Not used in Project.|
| _TaskID_|Optional|**Long**|Not used in Project.|
| _ProjectName_|Optional|**String**|Not used in Project.|

### Return Value

 **Boolean**


## Example

The following example sets the selected tasks in the Team Planner view to 40% complete. 


```vb
Sub TestSetTPField() 
    Dim fieldName As String 
 
    fieldName = FieldConstantToFieldName(pjTaskPercentComplete) 
    Application.SetTPField Field:=fieldName, Value:="40" 
End Sub
```


