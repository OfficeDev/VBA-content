---
title: Application.WorkbookNewChart Event (Excel)
keywords: vbaxl10.chm504115
f1_keywords:
- vbaxl10.chm504115
ms.prod: excel
api_name:
- Excel.Application.WorkbookNewChart
ms.assetid: 8456e472-6ea5-a916-10d6-f12afefb58fc
ms.date: 06/08/2017
---


# Application.WorkbookNewChart Event (Excel)

Occurs when a new chart is created in any open workbook.


## Syntax

 _expression_ . **WorkbookNewChart**( **_Wb_** , **_Ch_** )

 _expression_ A variable that represents an **[Application](application-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Wb_|Required| **[Workbook](workbook-object-excel.md)**|The workbook.|
| _Ch_|Required| **[Chart](chart-object-excel.md)**|The new chart.|

### Return Value

Nothing


## Remarks

The  **WorkbookNewChart** event occurs when a new chart is inserted or pasted in a worksheet, a chart sheet, or other sheet types. If multiple charts are inserted or pasted, the event will occur for each chart in the order they are inserted or pasted. If a chart object or chart sheet is moved from one location to another, the event will not occur. However, if the chart is moved between a chart object and a chart sheet, the event will occur because a new chart must be created.

The  **WorkbookNewChart** event will not occur in the following scenarios: copying or pasting a chart sheet, changing a chart type, changing a chart data source, undoing or redoing inserting or pasting a chart, and loading a workbook that contains a chart.


## Example

The following code example displays a message box when a new chart is added to a workbook.


```vb
Private Sub App_NewChart(ByVal Wb As Workbook, _ 
 ByVal Ch As Chart) 
 MsgBox ("A new chart was added to: " &; Wb.Name &; " of type: " &; Ch.Type) 
End Sub
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

