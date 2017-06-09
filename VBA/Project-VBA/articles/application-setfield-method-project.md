---
title: Application.SetField Method (Project)
keywords: vbapj.chm3
f1_keywords:
- vbapj.chm3
ms.prod: project-server
api_name:
- Project.Application.SetField
ms.assetid: 9f0670a9-b7e3-0bb6-40fc-0dcae63a3c19
ms.date: 06/08/2017
---


# Application.SetField Method (Project)

Sets the value of a local custom field or enterprise custom field for the selected tasks or resources.


## Syntax

 _expression_. **SetField**( ** _Field_**, ** _Value_**, ** _Create_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Field_|Required|**String**| The name of the field to set.|
| _Value_|Required|**String**|The value of the field.|
| _Create_|Optional|**Boolean**|**True** if a new value is created; otherwise **false**. The default value is **true**.|

### Return Value

 **Boolean**


## Remarks

If the custom field uses a lookup table that does not allow additional items to be entered, the specified Value must match a predefined value in the lookup table.

If the value of the Field argument does not exist as a custom field name for the seleted items, the  **SetField** method results in run-time error 1101.


## Example

The following example sets the value of an enterprise task text custom field to one of the valid values in the lookup table for the custom field. To use the example, create a lookup table in Project Web App that includes the value  **Value 3**, and then create a task custom text field that uses that lookup table. Select a task in the active project and run the command in the  **Immediate** window in the Visual Basic Editor.


```vb
Application.SetField Field:="TestEntTaskText", Value:="Value 3"
```


