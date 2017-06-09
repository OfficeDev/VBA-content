---
title: Application.SheetPivotTableUpdate Event (Excel)
keywords: vbaxl10.chm504094
f1_keywords:
- vbaxl10.chm504094
ms.prod: excel
api_name:
- Excel.Application.SheetPivotTableUpdate
ms.assetid: f42d1f7b-6395-326b-4b4f-72b497c81bd3
ms.date: 06/08/2017
---


# Application.SheetPivotTableUpdate Event (Excel)

Occurs after the sheet of the PivotTable report has been updated.


## Syntax

 _expression_ . **SheetPivotTableUpdate**( **_Sh_** , **_Target_** )

 _expression_ An expression that returns a **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Sh_|Required| **Object**|The selected sheet.|
| _Target_|Required| **PivotTable**|The selected PivotTable report.|

## Example

This example displays a message stating that the sheet of the PivotTable report has been updated. This example assumes you have declared an object of type  **[Application](application-object-excel.md)** or **[Workbook](workbook-object-excel.md)** with events in a class module.


```vb
Private Sub ConnectionApp_SheetPivotTableUpdate(ByVal shOne As Object, Target As PivotTable) 
 
 MsgBox "The SheetPivotTable connection has been updated." 
 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

