---
title: Chart.Select Method (Project)
keywords: vbapj.chm131638
f1_keywords:
- vbapj.chm131638
ms.prod: project-server
ms.assetid: dd4e1adf-3098-61a3-5913-8debc7d01351
ms.date: 06/08/2017
---


# Chart.Select Method (Project)
Selects one or more charts in a report.

## Syntax

 _expression_. **Select** _(Replace)_

 _expression_ A variable that represents a **Chart** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Replace_|Optional|**Variant**|**True** to replace the current selection with the specified chart. **False** to extend the current selection to include any previously selected charts.|
| _Replace_|Optional|VARIANT||

### Return value

 **Variant**


## Example

Create a report that contains at least two charts. The following example selects both charts in the report.


```vb
Sub SelectBothCharts()
    Dim chartShape1 As Shape
    Dim chartShape2 As Shape
    Dim reportName As String
    
    reportName = "Simple scalar chart"
    Set chartShape1 = ActiveProject.Reports(reportName).Shapes(1)
    Set chartShape2 = ActiveProject.Reports(reportName).Shapes(2)
    
    chartShape1.Chart.Select Replace:=True
    chartShape2.Chart.Select Replace:=False
End Sub
```


## See also


#### Other resources


[Chart Object](chart-object-project.md)
