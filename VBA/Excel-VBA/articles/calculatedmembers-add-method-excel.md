---
title: CalculatedMembers.Add Method (Excel)
keywords: vbaxl10.chm684078
f1_keywords:
- vbaxl10.chm684078
ms.prod: excel
api_name:
- Excel.CalculatedMembers.Add
ms.assetid: 8c6591bb-3906-6682-4dc7-89ffc2ae74f3
ms.date: 06/08/2017
---


# CalculatedMembers.Add Method (Excel)

Adds a calculated field or calculated item to a PivotTable. Returns a  **[CalculatedMember](calculatedmember-object-excel.md)** object.


## Syntax

 _expression_ . **Add**( **_Name_** , **_Formula_** , **_SolveOrder_** , **_Type_** )

 _expression_ A variable that represents a **CalculatedMembers** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the calculated member.|
| _Formula_|Required| **String**|The formula of the calculated member.|
| _SolveOrder_|Optional| **Variant**|The solve order for the calculated member.|
| _Type_|Optional| **Variant**|The type of calculated member.|
| _Dynamic_|Optional| **Boolean**|Specifies if the calculated member is recalculated with every update.|
| _DisplayFolder_|Optional| **String**|The name of the display folder for the calculated member.|
| _HierarchizeDistinct_|Optional| **Boolean**|Specifies whether to order and remove duplicates when displaying the hierarchy of the calculated member in a PivotTable report based on an OLAP cube.|

### Return Value

A  **CalculatedMember** object that represents the new calculated field or calculated item.


## Remarks

The  _Formula_ argument must contain a valid MDX (Multidimensional Expression) syntax statement. The _Name_ argument has to be acceptable to the Online Analytical Processing (OLAP) provider and the _Type_ argument has to be defined.

If you set the  _Type_ argument of this method to **xlCalculatedSet** , then you must call the **[AddSet](cubefields-addset-method-excel.md)** method to make the new field set visible in the PivotTable.


## Example

The following example adds a set to a PivotTable.


 **Note**  Connection to the cube and existing pivot table is necessary for the sample to run.


```vb
Sub UseAddSet() 
 
 Dim pvtOne As PivotTable 
 Dim strAdd As String 
 Dim strFormula As String 
 Dim cbfOne As CubeField 
 
 Set pvtOne = ActiveSheet.PivotTables(1) 
 
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


[CalculatedMembers Collection](calculatedmembers-object-excel.md)

