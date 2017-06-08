---
title: Charts.Add2 Method (Excel)
keywords: vbaxl10.chm218076
f1_keywords:
- vbaxl10.chm218076
ms.prod: excel
ms.assetid: bfd7d614-a640-dfdc-ebc5-3d0682f2c839
ms.date: 06/08/2017
---


# Charts.Add2 Method (Excel)

Inserts a chart directly onto the grid.


## Syntax

 _expression_ . **Add2**_(Before,_ _After,_ _Count,_ _NewLayout)_

 _expression_ A variable that represents a **Charts** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Before_|Optional|VARIANT|An object that specifies the sheet before which the new sheet is added.|
| _After_|Optional|VARIANT|An object that specifies the sheet after which the new sheet is added.|
| _Count_|Optional|VARIANT|The number of sheets to be added. The default value is one.|
| _NewLayout_|Optional|VARIANT|If  **NewLayout** is **True** , the chart is inserted by using the new dynamic formatting rules (Title is on, and Legend is on only if there are multiple series).|

### Return value

 **CHART**


## See also


#### Concepts


[Charts Collection](charts-object-excel.md)

