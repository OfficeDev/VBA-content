---
title: Application.CheckField Method (Project)
keywords: vbapj.chm7
f1_keywords:
- vbapj.chm7
ms.prod: project-server
api_name:
- Project.Application.CheckField
ms.assetid: a3360541-faa7-169e-1b23-5b3937fc6c07
ms.date: 06/08/2017
---


# Application.CheckField Method (Project)

**True** if the selected tasks or resources meet the specified criteria.


## Syntax

 _expression_. **CheckField** (**_Field_**, **_Value_**, **_Test_**, **_Op_**, **_Field2_**, **_Value2_**, **_Test2_**)

 _expression_ A variable that represents an **Application** object.


### Parameters

|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Field_|Required|**String**|The name of the field to search.|
| _Value_|Required|**String**|The value to compare with the value of the field specified with **Field**.|
| _Test_|Optional|**String**|The type of comparison made between **Field** and **Value**. The default value is "equals". Can be one of the following [comparison strings](#comparison-strings).|
| _Op_|Optional|**String**|How the criteria established with **Field**, **Test**, and **Value** relate to the second criteria. The **Op** argument can be set to "And" or "Or".|
| _Field2_|Optional|**String**|The name of a second field to search.|
| _Value2_|Optional|**String**|The value to compare with the value of the field specified with **Field2**.|
| _Test2_|Optional|**String**|The type of comparison made between **Field2** and **Value2**. Can be one of the same [comparison strings](#comparison-strings) as **Test**.|

<br/>

#### Comparison strings

|**Comparison string**|**Description**|
|:-----|:-----|
|"equals"|The value of **Field** equals **Value**.|
|"does not equal"|The value of **Field** does not equal **Value**.|
|"is greater than"|The value of **Field** is greater than **Value**.|
|"is greater than or equal to"|The value of **Field** is greater than or equal to **Value**.|
|"is less than"|The value of **Field** is less than **Value**.|
|"is less than or equal to"|The value of **Field** is less than or equal to **Value**.|
|"is within"|The value of **Field** is within **Value**.|
|"is not within"|The value of **Field** is not within **Value**.|
|"contains"|**Field** contains **Value**.|
|"does not contain"|**Field** does not contain **Value**.|
|"contains exactly"|**Field** exactly contains **Value**.|

<br/>

### Return value

 **Variant**


## Example

The following example determines whether value of Duration is equal to 1 and displays an appropriate message.

```vb
Sub Check_Field() 
 
 Dim T As Task 
 Dim Result As Boolean 
 
 Set T = ActiveProject.Tasks(3) 
 Result = CheckField("Duration", "1", "equals") 
 
 If Result Then 
 Result = MsgBox(T.GetField(pjTaskName) + " task Duration is equal to value specified.", vbOKOnly, "CheckField Method") 
 Else 
 Result = MsgBox(T.GetField(pjTaskName) + " task Duration is not equal to value specified.", vbOKOnly, "CheckField Method") 
 End If 
End Sub
```


