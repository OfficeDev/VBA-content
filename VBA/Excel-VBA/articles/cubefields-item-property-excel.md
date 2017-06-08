---
title: CubeFields.Item Property (Excel)
keywords: vbaxl10.chm670074
f1_keywords:
- vbaxl10.chm670074
ms.prod: excel
api_name:
- Excel.CubeFields.Item
ms.assetid: d068ccda-13e0-9938-7945-e8639e79d089
ms.date: 06/08/2017
---


# CubeFields.Item Property (Excel)

Returns a single object from a collection.


## Syntax

 _expression_ . **Item**( **_Index_** )

 _expression_ A variable that represents a **CubeFields** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number of the object.|

## Example

This example finds the first PivotTable report whose first cube field name contains the string ?Paris?. The  **Boolean** variable `blnFoundName` is set to **True** if the name is found.


```vb
blnFoundName = False 
For Each objPT in ActiveSheet.PivotTables 
 Set objCubeField = _ 
 objPT.CubeFields.Item(1) 
 If instr(1,objCubeField.Name, "Paris") <> 0 Then 
 blnFoundName = True 
 Exit For 
 End If 
Next objPT
```


## See also


#### Concepts


[CubeFields Object](cubefields-object-excel.md)

