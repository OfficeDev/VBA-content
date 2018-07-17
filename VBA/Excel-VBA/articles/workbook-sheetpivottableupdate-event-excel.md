---
title: Workbook.SheetPivotTableUpdate Event (Excel)
keywords: vbaxl10.chm503093
f1_keywords:
- vbaxl10.chm503093
ms.prod: excel
api_name:
- Excel.Workbook.SheetPivotTableUpdate
ms.assetid: 0b37939a-28dd-ef8b-ea5e-fc3768f8979a
ms.date: 06/08/2017
---


# Workbook.SheetPivotTableUpdate Event (Excel)

Occurs after the sheet of the PivotTable report has been updated.


## Syntax

 _expression_ . **SheetPivotTableUpdate**( **_Sh_** , **_Target_** )

 _expression_ An expression that returns a **Workbook** object.


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


[Workbook Object](workbook-object-excel.md)

