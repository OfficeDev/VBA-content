---
title: Scenarios.Add Method (Excel)
keywords: vbaxl10.chm362073
f1_keywords:
- vbaxl10.chm362073
ms.prod: excel
api_name:
- Excel.Scenarios.Add
ms.assetid: 0f76a5fd-82f1-7fa0-34f7-733b0e964666
ms.date: 06/08/2017
---


# Scenarios.Add Method (Excel)

Creates a new scenario and adds it to the list of scenarios that are available for the current worksheet.


## Syntax

 _expression_ . **Add**( **_Name_** , **_ChangingCells_** , **_Values_** , **_Comment_** , **_Locked_** , **_Hidden_** )

 _expression_ A variable that represents a **Scenarios** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The scenario name.|
| _ChangingCells_|Required| **Variant**|A  **[Range](range-object-excel.md)** object that refers to the changing cells for the scenario.|
| _Values_|Optional| **Variant**|An array that contains the scenario values for the cells in  _ChangingCells_. If this argument is omitted, the scenario values are assumed to be the current values in the cells in _ChangingCells_.|
| _Comment_|Optional| **Variant**|A string that specifies comment text for the scenario. When a new scenario is added, the author's name and date are automatically added at the beginning of the comment text.|
| _Locked_|Optional| **Variant**| **True** to lock the scenario to prevent changes. The default value is **True** .|
| _Hidden_|Optional| **Variant**| **True** to hide the scenario. The default value is **False** .|

### Return Value

A  **[Scenario](scenario-object-excel.md)** object that represents the new scenario.


## Remarks

A scenario name must be unique; Microsoft Excel generates an error if you try to create a scenario with a name that's already in use.


## Example

This example adds a new scenario to Sheet1.


```vb
Worksheets("Sheet1").Scenarios.Add Name:="Best Case", _ 
 ChangingCells:=Worksheets("Sheet1").Range("A1:A4"), _ 
 Values:=Array(23, 5, 6, 21), _ 
 Comment:="Most favorable outcome."
```


## See also


#### Concepts


[Scenarios Object](scenarios-object-excel.md)

