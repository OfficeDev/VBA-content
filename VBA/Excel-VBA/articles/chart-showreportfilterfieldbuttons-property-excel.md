---
title: Chart.ShowReportFilterFieldButtons Property (Excel)
keywords: vbaxl10.chm149189
f1_keywords:
- vbaxl10.chm149189
ms.prod: excel
api_name:
- Excel.Chart.ShowReportFilterFieldButtons
ms.assetid: 6b7aa6e2-2216-caef-5936-d9c9681b60db
ms.date: 06/08/2017
---


# Chart.ShowReportFilterFieldButtons Property (Excel)

Returns or sets whether to display the report filter field buttons on a PivotChart. Read/write


## Syntax

 _expression_ . **ShowReportFilterFieldButtons**

 _expression_ A variable that represents a **[Chart](chart-object-excel.md)** object.


### Return Value

 **Boolean**


## Remarks

Set the  **ShowReportFilterFieldButtons** property to **True** to display the **Report Filter Field** buttons on the specified PivotChart. Set the property to **False** to hide the buttons.

The  **ShowReportFilterFieldButtons** property corresponds to the **Show Report Filter Field Buttons** command on the **Field Buttons** drop-down list of the **Analyze** tab, which is available when a PivotChart is selected.


## Example

The following code example sets Chart 1 to display report filter field buttons.


```vb
ActiveSheet.ChartObjects("Chart 1").Activate 
ActiveChart.ShowReportFilterFieldButtons = True
```


## See also


#### Concepts


[Chart Object](chart-object-excel.md)

