---
title: Series.Paste Method (Excel)
keywords: vbaxl10.chm578100
f1_keywords:
- vbaxl10.chm578100
ms.prod: excel
api_name:
- Excel.Series.Paste
ms.assetid: 73e689cb-b2aa-61d7-e84c-113091d09a44
ms.date: 06/08/2017
---


# Series.Paste Method (Excel)

Pastes a picture from the Clipboard as the marker on the selected series.


## Syntax

 _expression_ . **Paste**

 _expression_ A variable that represents a **Series** object.


### Return Value

Variant


## Remarks

This method can be used on column, bar, line, or radar charts, and it sets the  **[MarkerStyle](series-markerstyle-property-excel.md)** property to **xlMarkerStylePicture** .


## Example

This example pastes a picture from the Clipboard into series one in Chart1.


```vb
Charts("Chart1").SeriesCollection(1).Paste
```


## See also


#### Concepts


[Series Object](series-object-excel.md)

