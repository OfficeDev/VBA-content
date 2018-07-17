---
title: Application.SetMatchingField Method (Project)
keywords: vbapj.chm11
f1_keywords:
- vbapj.chm11
ms.prod: project-server
api_name:
- Project.Application.SetMatchingField
ms.assetid: fcd57c26-6463-8821-481f-0c38d072118a
ms.date: 06/08/2017
---


# Application.SetMatchingField Method (Project)

Sets the value in the field of selected tasks or resources that meet the specified criteria.


## Syntax

_expression_. **SetMatchingField** (**_Field_**, **_Value_**, **_CheckField_**, **_CheckValue_**, **_CheckTest_**, **_CheckOperation_**, **_CheckField2_**, **_CheckValue2_**, **_CheckTest2_**)

_expression_ A variable that represents an **Application** object.


### Parameters

|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Field_|Required|**String**|The name of the field to set.|
| _Value_|Required|**String**|The value to which the field is set.|
| _CheckField_|Required|**String**|The name of the field to check.|
| _CheckValue_|Required|**String**|The value to compare with the value of the field specified with _CheckField_.|
| _CheckTest_|Optional|**String**|The type of comparison made between _CheckField_ and _CheckValue_. The default value is "equals". Can be one of the [comparison strings](#comparison-strings).|
| _CheckOperation_|Optional|**String**|How the criteria established with _CheckField_, _CheckTest_, and _CheckValue_ relate to the second criteria, if specified. The _CheckOperation_ argument can be set to "And" or "Or". The default value is "And".|
| _CheckField2_|Required|**String**|The name of the second field to check.|
| _CheckValue2_|Required|**String**|The value to which the second field is set.|
| _CheckTest2_|Optional|**Variant**|The type of comparison made between _CheckField2_ and  _CheckValue2_. Can be one of the same comparison strings as _CheckTest_.|

<br/>

#### Comparison strings

|**Comparison string**|**Description**|
|:-----|:-----|
|"equals"|The value of _CheckField_ equals _CheckValue_.|
|"does not equal"|The value of _CheckField_ does not equal _CheckValue_.|
|"is greater than"|The value of _CheckField_ is greater than _CheckValue_.|
|"is greater than or equal to"|The value of _CheckField_ is greater than or equal to _CheckValue_.|
|"is less than"|The value of _CheckField_ is less than _CheckValue_.|
|"is less than or equal to"|The value of _CheckField_ is less than or equal to _CheckValue_.|
|"is within"|The value of _CheckField_ is within _CheckValue_.|
|"is not within"|The value of _CheckField_ is not within _CheckValue_.|
|"contains"|_CheckField_ contains _CheckValue_.|
|"does not contain"|_CheckField_ does not contain _CheckValue_.|
|"contains exactly"|_CheckField_ exactly contains _CheckValue_.|

<br/>

### Return value

 **Boolean**

## Example

The following example checks for equality of task field `Name`, changes the value to `New Task Name`, and then changes the name back to the original.


```vb
Sub Set_MatchingField() 
 
 Dim T As Task 
 Dim OldName As String 
 
 'Save the task name 
 Set T = ActiveProject.Tasks(3) 
 OldName = T.GetField(pjTaskName) 
 
 ViewApply Name:="&;Gantt Chart" 
 'Change the field to "New Task's Name" 
 SetMatchingField Field:="Name", Value:="New Task Name", CheckField:="Name", CheckValue:=OldName, CheckTest:="equals" 
 ' Set the field to the old name 
 SetMatchingField Field:="Name", Value:=OldName, CheckField:="Name", CheckValue:="New Task's Name", CheckTest:="equals" 
End Sub
```


