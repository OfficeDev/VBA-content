---
title: Worksheets.Add2 Method (Excel)
keywords: vbaxl10.chm470090
f1_keywords:
- vbaxl10.chm470090
ms.prod: excel
ms.assetid: 4ae91335-f714-45e4-9677-6dfece31342e
ms.date: 06/08/2017
---


# Worksheets.Add2 Method (Excel)

This method is only implemented for the  **Charts** collection object and will produce a run time error if used on the **Sheets** and **Worksheets** objects.


## Syntax

 _expression_ . **Add2**_(Before,_ _After,_ _Count,_ _NewLayout)_

 _expression_ A variable that represents a **Worksheets** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Before_|Optional|VARIANT|An object that specifies the sheet before which the new sheet is added.|
| _After_|Optional|VARIANT|An object that specifies the sheet after which the new sheet is added.|
| _Count_|Optional|VARIANT|The number of sheets to be added. The default value is one.|
| _NewLayout_|Optional|VARIANT|The layout of the new worksheet.|

### Return value

 **OBJECT**


## See also


#### Concepts


[Worksheets Object](worksheets-object-excel.md)

