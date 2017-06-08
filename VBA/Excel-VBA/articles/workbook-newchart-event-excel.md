---
title: Workbook.NewChart Event (Excel)
keywords: vbaxl10.chm503108
f1_keywords:
- vbaxl10.chm503108
ms.prod: excel
api_name:
- Excel.Workbook.NewChart
ms.assetid: 76e7f325-9244-fd8c-b38d-063f0193a5e9
ms.date: 06/08/2017
---


# Workbook.NewChart Event (Excel)

Occurs when a new chart is created in the workbook.


## Syntax

 _expression_ . **NewChart**( **_Ch_** )

 _expression_ A variable that returns a **[Workbook](workbook-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Ch_|Required| **[Chart](chart-object-excel.md)**|The new chart.|

### Return Value

Nothing


## Remarks

The  **NewChart** event occurs whenever a new chart is inserted or pasted in a worksheet, a chart sheet, or other sheet types. If multiple charts are inserted or pasted, the event will occur for each chart in the order they are inserted or pasted. If a chart object or chart sheet is moved from one location to another, the event will not occur. However, if the chart is moved between a chart object and a chart sheet, the event will occur because a new chart must be created.

The  **NewChart** event will not occur in the following scenarios: copying or pasting a chart sheet, changing a chart type, changing a chart data source, undoing or redoing inserting or pasting a chart, and loading a workbook that contains a chart.


## Example

This example displays a message box when a new chart is added to the workbook.


```vb
Private Sub Workbook_NewChart(ByVal Ch As Chart) 
 MsgBox ("A new chart of the following chart type was added: " &; Ch.ChartType) 
End Sub
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

