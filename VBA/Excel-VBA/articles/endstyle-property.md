---
title: EndStyle Property
ms.prod: excel
api_name:
- Excel.EndStyle
ms.assetid: 2d12c0c5-7c48-41c0-b270-d5cf70eb7d47
ms.date: 06/08/2017
---


# EndStyle Property

Returns or sets the end style for the error bars. Read/write  **XlEndStyleCap**.


## 



|XlEndStyleCap can be one of these XlEndStyleCap constants.|
| **xlCap**|
| **xlNoCap**|
 _expression_.  **EndStyle**

 _expression_ Required. An expression that returns one of the objects in the Applies To list.


## Example

This example sets the end style for the error bars for series one. The example should be run on a 2-D line chart that has Y error bars for the first series.


```
myChart.SeriesCollection(1).ErrorBars. EndStyle = xlCap

```


