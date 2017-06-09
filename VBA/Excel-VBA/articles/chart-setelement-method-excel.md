---
title: Chart.SetElement Method (Excel)
keywords: vbaxl10.chm149175
f1_keywords:
- vbaxl10.chm149175
ms.prod: excel
api_name:
- Excel.Chart.SetElement
ms.assetid: 0efff437-179b-fe16-118b-6f3cde49c5cf
ms.date: 06/08/2017
---


# Chart.SetElement Method (Excel)

Sets chart elements on a chart. Read/write  **MsoChartElementType** .


## Syntax

 _expression_ . **SetElement**( **_Element_** )

 _expression_ A variable that represents a **Chart** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Element_|Required| **MsoChartElementType**|Specifies the chart element type.|

### Return Value

Nothing


## Remarks

For charts, the following commands in the  **Layout** tab correspond to the **SetElement** method:


- Everything in the  **Labels** group.
    
- Everything in the  **Axes** group.
    
- Everything in the  **Analysis** group.
    
-  **PlotArea**,  **Chart Wall**, and  **Chart Floor** buttons.
    


 **MsoChartElementType** is an enumeration of constants that refer to all of the above commands.


## Example

This example sets chart elements using the various constant values to an active chart.


```vb
ActiveChart.Axes(xlValue).MajorGridlines.Select 
 ActiveChart.SetElement (msoElementChartTitleCenteredOverlay) 
 ActiveChart.SetElement (msoElementPrimaryCategoryGridLinesMinor) 
 ActiveChart.Walls.Select 
 Application.CommandBars("Clip Art").Visible = False 
 ActiveChart.SetElement (msoElementChartFloorShow)
```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

