---
title: ServerViewableItems.Delete Method (Excel)
keywords: vbaxl10.chm833075
f1_keywords:
- vbaxl10.chm833075
ms.prod: excel
api_name:
- Excel.ServerViewableItems.Delete
ms.assetid: e6b53271-8a37-4bf3-fea2-46d02550391b
ms.date: 06/08/2017
---


# ServerViewableItems.Delete Method (Excel)

Deletes a reference to an object in the  **[ServerViewableItems](serverviewableitems-object-excel.md)** collection in the workbook.


## Syntax

 _expression_ . **Delete**( **_Index_** )

 _expression_ A variable that represents a **ServerViewableItems** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The index of the object you want to delete.|

## Remarks

If you do not want a particular object to be viewable in Excel Services, use this method to remove that object.


## See also


#### Concepts


[ServerViewableItems Object](serverviewableitems-object-excel.md)

