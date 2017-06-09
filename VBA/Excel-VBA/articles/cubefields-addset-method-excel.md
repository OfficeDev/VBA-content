---
title: CubeFields.AddSet Method (Excel)
keywords: vbaxl10.chm670077
f1_keywords:
- vbaxl10.chm670077
ms.prod: excel
api_name:
- Excel.CubeFields.AddSet
ms.assetid: 2f40d4f3-56fc-4d98-b214-623885dc26d6
ms.date: 06/08/2017
---


# CubeFields.AddSet Method (Excel)

Adds a new  **[CubeField](cubefield-object-excel.md)** object to the **[CubeFields](cubefields-object-excel.md)** collection. The **CubeField** object corresponds to a set defined on the Online Analytical Processing (OLAP) provider for the cube.


## Syntax

 _expression_ . **AddSet**( **_Name_** , **_Caption_** )

 _expression_ A variable that represents a **CubeFields** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|A valid name in the SETS schema rowset.|
| _Caption_|Required| **String**|A string representing the field that will be displayed in the PivotTable view.|

### Return Value

CubeField


## Remarks

If a set with the name given in the argument  _Name_ does not exist, the **AddSet** method will return a run-time error.


## Example

In this example, Microsoft Excel adds a set titled "My Set" to the  **CubeField** object. This example assumes an OLAP PivotTable report exists on the active worksheet. Also, this example assumes a field titled "Product" exists.


```vb
Sub UseAddSet() 
 
 Dim pvtOne As PivotTable 
 Dim strAdd As String 
 Dim strFormula As String 
 Dim cbfOne As CubeField 
 
 Set pvtOne = Sheet1.PivotTables(1) 
 
 strAdd = "[MySet]" 
 strFormula = "'{[Product].[All Products].[Food].children}'" 
 
 ' Establish connection with data source if necessary. 
 If Not pvtOne.PivotCache.IsConnected Then pvtOne.PivotCache.MakeConnection 
 
 ' Add a calculated member titled "[MySet]" 
 pvtOne.CalculatedMembers.Add Name:=strAdd, _ 
 Formula:=strFormula, Type:=xlCalculatedSet 
 
 ' Add a set to the CubeField object. 
 Set cbfOne = pvtOne.CubeFields.AddSet(Name:="[MySet]", _ 
 Caption:="My Set") 
 
End Sub
```


## See also


#### Concepts


[CubeFields Object](cubefields-object-excel.md)

