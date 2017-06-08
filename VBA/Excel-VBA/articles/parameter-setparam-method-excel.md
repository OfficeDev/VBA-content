---
title: Parameter.SetParam Method (Excel)
keywords: vbaxl10.chm523079
f1_keywords:
- vbaxl10.chm523079
ms.prod: excel
api_name:
- Excel.Parameter.SetParam
ms.assetid: af1f5b0a-75a1-ae85-b291-cc3ab514b0a3
ms.date: 06/08/2017
---


# Parameter.SetParam Method (Excel)

Defines a parameter for the specified query table.


## Syntax

 _expression_ . **SetParam**( **_Type_** , **_Value_** )

 _expression_ A variable that represents a **Parameter** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **[XlParameterType](xlparametertype-enumeration-excel.md)**|One of the constants of  **XlParameterType** which specifies the parameter type.|
| _Value_|Required| **Variant**|The value of the specified parameter, as shown in the description of the  _Type_ argument.|

## Remarks





| **XlParameterType** can be one of these **XlParameterType** constants.|
| **xlConstant** . Uses the value specified by the _Value_ argument.|
| **xlPrompt** . Displays a dialog box that prompts the user for the value. The _Value_ argument specifies the text shown in the dialog box.|
| **xlRange** . Uses the value of the cell in the upper-left corner of the range. The _Value_ argument specifies a **[Range](range-object-excel.md)** object|

## Example

This example changes the SQL statement for query table one. The clause ?(city=?)? indicates that the query is a parameter query, and the example sets the value of city to the constant ?Oakland.?


```sql
Set qt = Sheets("sheet1").QueryTables(1) 
qt.Sql = "SELECT * FROM authors WHERE (city=?)" 
Set param1 = qt.Parameters.Add("City Parameter", _ 
 xlParamTypeVarChar) 
param1.SetParam xlConstant, "Oakland" 
qt.Refresh
```


## See also


#### Concepts


[Parameter Object](parameter-object-excel.md)

