---
title: Chart.ShowLegendFieldButtons Property (Excel)
keywords: vbaxl10.chm149190
f1_keywords:
- vbaxl10.chm149190
ms.prod: excel
api_name:
- Excel.Chart.ShowLegendFieldButtons
ms.assetid: 44f1554c-145b-8600-07c4-40b6891dab2d
ms.date: 06/08/2017
---


# Chart.ShowLegendFieldButtons Property (Excel)

Returns or sets whether to display legend field buttons on a PivotChart. Read/write


## Syntax

 _expression_ . **ShowLegendFieldButtons**

 _expression_ A variable that represents a **[Chart](chart-object-excel.md)** object.


### Return Value

 **Boolean**


## Remarks

Set the  **ShowLegendFieldButtons** property to **True** to display legend field buttons on the specified PivotChart. Set the property to **False** to hide the buttons.

The  **ShowLegendFieldButtons** property corresponds to the **Show Legend Field Buttons** command on the **Field Buttons** drop-down list of the **Analyze** tab, which is available when a PivotChart is selected.


## Example

The following code example sets Chart 1 to display legend field buttons.


```vb
ActiveSheet.ChartObjects("Chart 1").Activate 
ActiveChart.ShowLegendFieldButtons = True
```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

