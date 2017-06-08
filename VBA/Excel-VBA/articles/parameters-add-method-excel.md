---
title: Parameters.Add Method (Excel)
keywords: vbaxl10.chm525073
f1_keywords:
- vbaxl10.chm525073
ms.prod: excel
api_name:
- Excel.Parameters.Add
ms.assetid: 043276ed-4af7-3b7a-dbfb-549489d3a127
ms.date: 06/08/2017
---


# Parameters.Add Method (Excel)

Creates a new query parameter.


## Syntax

 _expression_ . **Add**( **_Name_** , **_iDataType_** )

 _expression_ A variable that represents a **Parameters** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the specified parameter. The parameter name should match the parameter clause in the SQL statement.|
| _iDataType_|Optional| **Variant**|The data type of the parameter. Can be any  **[XlParameterDataType](xlparameterdatatype-enumeration-excel.md)** constant. These values correspond to ODBC data types. They indicate the type of value the ODBC driver is expecting to receive. Microsoft Excel and the ODBC driver manager will coerce the parameter value given in Microsoft Excel into the correct data type for the driver.|

### Return Value

A  **[Parameter](parameter-object-excel.md)** object that represents the new query parameter.


## Example

This example changes the SQL statement for query table one. The clause "(city=?)" indicates that the query is a parameter query, and the value of city is set to the constant "Oakland."


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


[Parameters Object](parameters-object-excel.md)

