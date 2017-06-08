---
title: Chart.ShowAllFieldButtons Property (Excel)
keywords: vbaxl10.chm149193
f1_keywords:
- vbaxl10.chm149193
ms.prod: excel
api_name:
- Excel.Chart.ShowAllFieldButtons
ms.assetid: b5a9dc1a-2c85-eece-b678-2d3509780a46
ms.date: 06/08/2017
---


# Chart.ShowAllFieldButtons Property (Excel)

Returns or sets whether to display all field buttons on a PivotChart. Read/write


## Syntax

 _expression_ . **ShowAllFieldButtons**

 _expression_ A variable that represents a **[Chart](chart-object-excel.md)** object.


### Return Value

 **Boolean**


## Remarks

Set the  **ShowAllFieldButtons** property to **True** to display all field buttons on the specified PivotChart. Set the property to **False** to hide all field buttons.

The  **ShowAllFieldButtons** property corresponds to the **Hide All** command on the **Field Buttons** drop-down list of the **Analyze** tab, which is available when a PivotChart is selected.


## Example

The following code example sets Chart 1 to display all field buttons.


```vb
ActiveSheet.ChartObjects("Chart 1").Activate 
ActiveChart.ShowAllFieldButtons = True
```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

