---
title: Chart.SetSourceData Method (Project)
keywords: vbapj.chm131632
f1_keywords:
- vbapj.chm131632
ms.prod: project-server
ms.assetid: 723680bb-f2ec-3a8f-f392-a6c90eae7ff8
ms.date: 06/08/2017
---


# Chart.SetSourceData Method (Project)
Sets a source data range from Excel for a chart.

## Syntax

 _expression_. **SetSourceData** _(Source,_? _PlotBy)_

 _expression_ A variable that represents a **Chart** object.


### Parameters



|**Name**|**Required/Optional**|**Data type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Source_|Required|**String**|The source data range.|
| _PlotBy_|Optional|**Variant**|Specifies the way the data is plotted. Can be one of the following  **Office.XlRowCol** constants: **xlColumns** or **xlRows**.|
| _Source_|Required|STRING||
| _PlotBy_|Optional|VARIANT||

### Return value

 **Nothing**


## Remarks

A chart in a Project report can use a data range from Excel, if Project programmatically accesses an Excel worksheet. The charting object model in Project accepts range address strings for properties and methods that accept  **Range** objects in Excel. A range address string in Project is expressed differently than a range in Excel. For example, the _Source_ parameter can have a range value such as `"='Sheet1'!$A$1:$D$5"`. 


## See also


#### Other resources


[Chart Object](chart-object-project.md)
