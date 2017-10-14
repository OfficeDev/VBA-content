---
title: Workbook.SheetPivotTableChangeSync Event (Excel)
keywords: vbaxl10.chm503106
f1_keywords:
- vbaxl10.chm503106
ms.prod: excel
api_name:
- Excel.Workbook.SheetPivotTableChangeSync
ms.assetid: c280b935-3dbf-0666-b727-64d6b4ac7ebd
ms.date: 06/08/2017
---


# Workbook.SheetPivotTableChangeSync Event (Excel)

Occurs after changes to a PivotTable.


## Syntax

 _expression_ . **SheetPivotTableChangeSync**( **_Sh_** , **_Target_** )

 _expression_ A variable that represents a **[Workbook](workbook-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Sh_|Required| **Object**|The worksheet that contains the PivotTable.|
| _Target_|Required| **[PivotTable](pivottable-object-excel.md)**|The PivotTable that was changed.|

### Return Value

Nothing


## Remarks

The  **PivotTableChangeEvent** occurs during most changes to a PivotTable, so that you can write code to respond to user actions, such as clearing, grouping, or refreshing items in the PivotTable.


## Example

The following code example displays a message box that shows the name of the PivotTable the user changed. 


```vb
Private Sub Workbook_SheetPivotTableChangeSync(ByVal Sh As Target, Target As PivotTable) 
 
With Target 
 MsgBox "You performed an operation in the following PivotTable: " &; .Name &; " on " &; Sh.Name 
End With 
 
End Sub
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

