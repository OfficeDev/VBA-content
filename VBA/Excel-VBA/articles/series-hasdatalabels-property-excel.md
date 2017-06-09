---
title: Series.HasDataLabels Property (Excel)
keywords: vbaxl10.chm578088
f1_keywords:
- vbaxl10.chm578088
ms.prod: excel
api_name:
- Excel.Series.HasDataLabels
ms.assetid: 10f879c9-4d34-d20b-facc-44ebc950aaa2
ms.date: 06/08/2017
---


# Series.HasDataLabels Property (Excel)

 **True** if the series has data labels. Read/write **Boolean** .


## Syntax

 _expression_ . **HasDataLabels**

 _expression_ A variable that represents a **Series** object.


## Example

This example turns on data labels for series three in Chart1.


```vb
With Charts("Chart1").SeriesCollection(3) 
 .HasDataLabels = True 
 .ApplyDataLabels Type:=xlValue 
End With
```


## See also


#### Concepts


[Series Object](series-object-excel.md)

