---
title: DataLabel.ShowCategoryName Property (Excel)
keywords: vbaxl10.chm582100
f1_keywords:
- vbaxl10.chm582100
ms.prod: excel
api_name:
- Excel.DataLabel.ShowCategoryName
ms.assetid: a8f2fdad-273a-3a45-7396-9691109c25d4
ms.date: 06/08/2017
---


# DataLabel.ShowCategoryName Property (Excel)

 **True** to display the category name for the data labels on a chart. **False** to hide. Read/write **Boolean** .


## Syntax

 _expression_ . **ShowCategoryName**

 _expression_ An expression that returns a **DataLabel** object.


## Remarks

The chart must first be active before you can access the data labels programmatically or a run-time error will occur.


## Example

This example shows the category name for the data labels of the first series on the first chart. This example assumes a chart exists on the active worksheet.


```vb
Sub UseCategoryName() 
 
 ActiveSheet.ChartObjects(1).Activate 
 ActiveChart.SeriesCollection(1) _ 
 .DataLabels.ShowCategoryName = True 
 
End Sub
```


## See also


#### Concepts


[DataLabel Object](datalabel-object-excel.md)

