---
title: Chart.ShowAxisFieldButtons Property (Excel)
keywords: vbaxl10.chm149191
f1_keywords:
- vbaxl10.chm149191
ms.prod: excel
api_name:
- Excel.Chart.ShowAxisFieldButtons
ms.assetid: 05eff4ce-c06b-b866-b0d7-8733cb51605a
ms.date: 06/08/2017
---


# Chart.ShowAxisFieldButtons Property (Excel)

Returns or sets whether to display axis field buttons on a PivotChart. Read/write


## Syntax

 _expression_ . **ShowAxisFieldButtons**

 _expression_ A variable that represents a **[Chart](chart-object-excel.md)** object.


### Return Value

 **Boolean**


## Remarks

Set the  **ShowAxisFieldButtons** property to **True** to display axis field buttons on the specified PivotChart. Set the property to **False** to hide the buttons.

The  **ShowAxisFieldButtons** property corresponds to the **Show Axis Field Buttons** command on the **Field Buttons** drop-down list of the **Analyze** tab, which is available when a PivotChart is selected.


## Example

The following code example sets Chart 1 to hide axis field buttons.


```vb
ActiveSheet.ChartObjects("Chart 1").Activate 
ActiveChart.ShowAxisFieldButtons = False
```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

