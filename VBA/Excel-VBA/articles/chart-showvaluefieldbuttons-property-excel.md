---
title: Chart.ShowValueFieldButtons Property (Excel)
keywords: vbaxl10.chm149192
f1_keywords:
- vbaxl10.chm149192
ms.prod: excel
api_name:
- Excel.Chart.ShowValueFieldButtons
ms.assetid: 7997b313-ce87-95eb-3d1e-b9b7b6eda84b
ms.date: 06/08/2017
---


# Chart.ShowValueFieldButtons Property (Excel)

Returns or sets whether to display the value field buttons on a PivotChart. Read/write


## Syntax

 _expression_ . **ShowValueFieldButtons**

 _expression_ A variable that represents a **[Chart](chart-object-excel.md)** object.


### Return Value

 **Boolean**


## Remarks

Set the  **ShowValueFieldButtons** property to **True** to display the **Value Field** buttons on the specified PivotChart. Set the property to **False** to hide the button.

The  **ShowValueFieldButtons** property corresponds to the **Show Value Field Buttons** command on the **Field Buttons** drop-down list of the **Analyze** tab, which is available when a PivotChart is selected.


## Example

The following code example sets Chart 1 to display value field buttons.


```vb
ActiveSheet.ChartObjects("Chart 1").Activate 
ActiveChart.ShowValueFieldButtons = True
```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

