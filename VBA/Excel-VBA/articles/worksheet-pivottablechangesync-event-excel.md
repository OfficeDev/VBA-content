---
title: Worksheet.PivotTableChangeSync Event (Excel)
keywords: vbaxl10.chm502086
f1_keywords:
- vbaxl10.chm502086
ms.prod: excel
api_name:
- Excel.Worksheet.PivotTableChangeSync
ms.assetid: b8cd1e24-4986-d3d4-c37a-b2933c6a9d99
ms.date: 06/08/2017
---


# Worksheet.PivotTableChangeSync Event (Excel)

Occurs after changes to a PivotTable.


## Syntax

 _expression_ . **PivotTableChangeSync**( **_Target_** )

 _expression_ A variable that represents a **[Worksheet](worksheet-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Target_|Required| **[PivotTable](pivottable-object-excel.md)**|The PivotTable that was changed.|

### Return Value

Nothing


## Remarks

The  **PivotTableChangeEvent** occurs during most changes to a PivotTable, so that you can write code to respond to user actions, such as clearing, grouping, or refreshing items in the PivotTable.


## Example

The following code example displays a message box that shows the name of the PivotTable the user changed. 


```vb
Private Sub Worksheet_PivotTableChangeSync(ByVal Target As PivotTable) 
 
With Target 
 MsgBox "You performed an operation in the following PivotTable: " &; .Name 
End With 
 
End Sub
```


## See also


#### Concepts


[Worksheet Object](worksheet-object-excel.md)

