---
title: Workbook.PivotTableOpenConnection Event (Excel)
keywords: vbaxl10.chm503095
f1_keywords:
- vbaxl10.chm503095
ms.prod: excel
ms.assetid: b6ce12f7-7bc6-bfcc-33f4-2e8ea6e53bae
ms.date: 06/08/2017
---


# Workbook.PivotTableOpenConnection Event (Excel)

Occurs after a PivotTable report opens the connection to its data source.


## Syntax

 _expression_ . **PivotTableOpenConnection**( **_Target_** )

 _expression_ A variable that represents a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Target_|Required| **[PivotTable](pivottable-object-excel.md)**|The selected PivotTable report.|
| _Target_|Required|PIVOTTABLE||
|Name|Required/Optional|Data type|Description|

### Return Value

Nothing


## Example

This example displays a message stating that the PivotTable report's connection to its source has been opened. This example assumes you have declared an object of type  **Workbook** with events in a class module.


```vb
Private Sub ConnectionApp_PivotTableOpenConnection(ByVal Target As PivotTable) 
 
 MsgBox "The PivotTable connection has been opened." 
 
End Sub
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

