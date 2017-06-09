---
title: Application.SelectTimescaleRange Method (Project)
keywords: vbapj.chm954
f1_keywords:
- vbapj.chm954
ms.prod: project-server
api_name:
- Project.Application.SelectTimescaleRange
ms.assetid: 16a4bd12-7a60-c172-6a73-c3552b2baf4b
ms.date: 06/08/2017
---


# Application.SelectTimescaleRange Method (Project)

Selects one or more timescale data cells in a usage view.


## Syntax

 _expression_. **SelectTimescaleRange**( ** _Row_**, ** _StartTime_**, ** _Width_**, ** _Height_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Row_|Required|**Long**|The number of the row containing the cell to select.|
| _StartTime_|Required|**String**|A time (from the timescale) that functions as the starting point of the selection.|
| _Width_|Required|**Integer**| The number of columns to select.|
| _Height_|Required|**Long**|The number of rows to select.|

### Return Value

 **Boolean**


## Example

The following example selects a five-day range of timescale data cells for the specified row. It assumes the timescale has not been changed from the default setting. The  **SelectRow** method is not required for this example, but is included to make the result easier to read.


```vb
Sub SelectWeek() 
 Dim WhichRow As Integer, StartDate As Variant 
 
 WhichRow = InputBox("Start selection on which row?") 
 StartDate = InputBox("Enter the date for the start of a week: ") 
 
 SelectRow WhichRow, False 
 SelectTimescaleRange Row:=WhichRow, StartTime:=StartDate, Width:=5, Height:=1 
 
End Sub
```


